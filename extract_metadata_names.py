#!/usr/bin/env python3

from pathlib import Path
from argparse import ArgumentParser, FileType
from sys import stdout
from typing import List, Set
from logging import getLogger, CRITICAL, ERROR, WARNING, INFO, DEBUG, NOTSET, LogRecord
from unicodedata import category as unicodedata_category

from magic import Magic
from terminal_utils.progressor import Progressor, ProgressStatus, ProgressorLogHandler
from terminal_utils.colored_output import ColoredOutput, PrintColor

from metadata_extractor.extractors.office_open_xml import OfficeOpenXMLMetadataExtractor
from metadata_extractor.extractors.ole import OLEMetadataExtractor
from metadata_extractor.extractors.pdf import PDFMetadataExtractor
from metadata_extractor.extractors import MetadataExtractor

LOG = getLogger(__name__)


class UnsupportedMimetypeError(Exception):
    def __init__(self, observed_mimetype: str, file_path: str):
        super().__init__(f'Unsupported mimetype {observed_mimetype} for {file_path}')
        self.observed_mimetype = observed_mimetype
        self.file_path = file_path


def extract_metadata_names(sources: List[Path], recurse: bool, exit_on_unsupported_mimetype: bool = False) -> Set[str]:
    """
    Extract names from metadata of documents.

    :param sources: Paths of document files or directories of document files from which to extract names.
    :param recurse: Whether to recursively find document files in directories.
    :param exit_on_unsupported_mimetype: Exit when the mimetype of an input file is not supported.
    :return: Names extracted from metadata of document files referred to.
    """

    get_mimetype = Magic(mime=True).from_file

    paths = []
    for source in sources:
        if source.is_dir():
            paths.extend(path for path in (source.glob('**/*' if recurse else '*')) if not path.is_dir())
        else:
            paths.append(source)

    names: Set[str] = set()

    for i, path in enumerate(paths, start=1):
        LOG.debug(ProgressStatus(iteration=i-1, total=len(paths), prefix='Extracting metadata... '))

        mimetype = get_mimetype(str(path))

        try:
            metadata_extractor = MetadataExtractor.extractor_from_mime_type(mime_type=mimetype)
        except KeyError:
            error_message = f'Unsupported mimetype {mimetype} for file {path}'
            if exit_on_unsupported_mimetype:
                raise UnsupportedMimetypeError(observed_mimetype=mimetype, file_path=str(path))
            else:
                LOG.warning(error_message)
                continue
        except Exception as e:
            raise e

        metadata = metadata_extractor.from_path(path=path)
        if not metadata:
            continue

        if metadata_extractor is OfficeOpenXMLMetadataExtractor:
            core_metadata = metadata.get('core.xml')

            if creator := core_metadata.get('dc:creator'):
                names.add(creator)

            if last_modified_by := core_metadata.get('cp:lastmodifiedby'):
                names.add(last_modified_by)

        elif metadata_extractor is OLEMetadataExtractor:
            if author_bytes := metadata.get('author'):
                names.add(author_bytes.decode(encoding='latin-1'))

            if last_saved_by_bytes := metadata.get('last_saved_by'):
                names.add(last_saved_by_bytes.decode(encoding='latin-1'))

        elif metadata_extractor is PDFMetadataExtractor:
            try:
                author_bytes: bytes = metadata.get('Author', b'')
                if author_bytes.startswith(b'\xfe\xff'):
                    author_bytes = author_bytes[2:]
                # NOTE: Allegedly, latin-1 is the default encoding in PDFs.
                if name := author_bytes.decode(encoding='latin-1'):
                    names.add(name)
            except Exception as e:
                LOG.warning(e)

    return {
        ''.join(character for character in name if unicodedata_category(character) in {'Lu', 'Ll', 'Zs'}).strip()
        for name in names
    }


class MetadataNamesArgumentParser(ArgumentParser):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.add_argument(
            'sources',
            help='Sources.',
            nargs='+',
            type=Path,
            metavar='SOURCE'
        )

        self.add_argument(
            '-r', '--recurse',
            help='Recurse directories.',
            dest='recurse',
            action='store_true',
            default=False
        )

        self.add_argument(
            '-o', '--output',
            help='A path to which the output should be written.',
            dest='output_destination',
            type=FileType('w'),
            default=stdout
        )

        self.add_argument(
            '-w', '--ignore-warnings',
            help='Do not print warning messages.',
            dest='ignore_warnings',
            action='store_true'
        )

        self.add_argument(
            '-q', '--quiet',
            help='Do not print warning messages, error messages, or the progress bar.',
            dest='quiet',
            action='store_true'
        )

        self.add_argument(
            '-e', '--exit-on-unsupported-mimetype',
            help='Exit when the mimetype of an input file is not supported',
            dest='exit_on_unsupported_mimetype',
            action='store_true'
        )


class ExtractMetadataNamesProgressLogHandler(ProgressorLogHandler):

    def __init__(self, progressor: Progressor, ignore_warnings: bool = False, level=NOTSET):
        super().__init__(progressor=progressor, level=level)
        self._colored_output = ColoredOutput()
        self.ignore_warnings = ignore_warnings

    def emit(self, record: LogRecord):
        if isinstance(record.msg, ProgressStatus):
            # record.msg.prefix = f'{self._progress_prefix} '
            super().emit(record=record)
            return

        formatted_message: str = self.format(record=record)

        if record.levelno in {CRITICAL, ERROR}:
            self._progressor.print_message(
                message=self._colored_output.make_color_output(print_color=PrintColor.RED, message=formatted_message)
            )
        elif record.levelno == WARNING:
            if not self.ignore_warnings:
                self._progressor.print_message(
                    message=self._colored_output.make_color_output(print_color=PrintColor.YELLOW, message=formatted_message)
                )
        elif record.levelno == INFO:
            self._progressor.print_message(
                message=self._colored_output.make_color_output(print_color=PrintColor.GREEN, message=formatted_message)
            )
        elif record.levelno == DEBUG:
            self._progressor.print_message(
                message=self._colored_output.make_color_output(print_color=PrintColor.WHITE, message=formatted_message)
            )
        else:
            raise ValueError(f'Unknown log level: levelno={record.levelno}')


def main():
    args = MetadataNamesArgumentParser().parse_args()

    try:
        with Progressor() as progressor:
            if args.quiet:
                LOG.disabled = True
            else:
                handler = ExtractMetadataNamesProgressLogHandler(
                    progressor=progressor,
                    ignore_warnings=args.ignore_warnings
                )
                LOG.addHandler(handler)
                LOG.setLevel(level=DEBUG)

            metadata_names = extract_metadata_names(
                sources=args.sources,
                recurse=args.recurse,
                exit_on_unsupported_mimetype=args.exit_on_unsupported_mimetype
            )
    except KeyboardInterrupt:
        pass
    except UnsupportedMimetypeError as e:
        LOG.error(e)
    except:
        LOG.exception('Unexpected error')
    else:
        print('\n'.join(sorted(metadata_names)))


if __name__ == '__main__':
    main()
