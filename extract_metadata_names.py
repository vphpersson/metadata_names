#!/usr/bin/env python3

from pathlib import Path
from argparse import ArgumentParser, FileType
from sys import stdout
from typing import List, Set
from logging import getLogger
from unicodedata import category as unicodedata_category

from magic import Magic
from terminal_utils.Progressor import Progressor

from metadata_names.extractors.msxml import MSXMLMetadataExtractor
from metadata_names.extractors.ole import OLEMetadataExtractor
from metadata_names.extractors.pdf import PDFMetadataExtractor
from metadata_names.extractors import MetadataExtractor

LOG = getLogger(__name__)


def extract_metadata_names(sources: List[Path], recurse: bool) -> Set[str]:
    """
    Extract names from metadata of documents.

    :param sources: Paths of document files or directories of document files from which to extract names.
    :param recurse: Whether to recursively find document files in directories.
    :return: Names extracted from metadata of document files referred to.
    """

    get_mimetype = Magic(mime=True).from_file

    paths = []
    for source in sources:
        if source.is_dir():
            paths.extend(source.glob('**/*' if recurse else '*'))
        else:
            paths.append(source)

    names: Set[str] = set()

    with Progressor() as progressor:
        msg = 'Extracting metadata... '

        progressor.print_progress(iteration=0, total=len(paths) or 1, prefix=msg)

        for i, path in enumerate(paths, start=1):

            try:
                metadata_extractor = MetadataExtractor.extractor_from_mime_type(mime_type=get_mimetype(str(path)))
                metadata = metadata_extractor.from_path(path=path)
            except Exception as e:
                LOG.warning(e)
                continue

            if not metadata:
                continue

            if metadata_extractor is MSXMLMetadataExtractor:
                core_metadata = metadata.get('core')
                if not core_metadata:
                    continue

                if creator := core_metadata.get('creator'):
                    names.add(creator)

                if last_modified_by := core_metadata.get('lastModifiedBy'):
                    names.add(last_modified_by)

            elif metadata_extractor is OLEMetadataExtractor:
                summary_information_metadata = metadata.get('SummaryInformation')
                if not summary_information_metadata:
                    continue

                try:
                    if author_bytes := summary_information_metadata.get('author'):
                        names.add(author_bytes.decode(encoding='latin-1'))
                except Exception as e:
                    LOG.warning(e)

                try:
                    if last_saved_by_bytes := summary_information_metadata.get('last_saved_by'):
                        names.add(last_saved_by_bytes.decode(encoding='latin-1'))
                except Exception as e:
                    LOG.warning(e)

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

            progressor.print_progress(i, len(paths), prefix=msg)

    return {
        ''.join(character for character in name if unicodedata_category(character) in {'Lu', 'Ll', 'Zs'}).strip()
        for name in names
    }


def get_parser() -> ArgumentParser:
    parser = ArgumentParser()

    parser.add_argument(
        'sources',
        help='Sources.',
        nargs='+',
        type=Path,
        metavar='SOURCE'
    )

    parser.add_argument(
        '-r', '--recurse',
        help='Recurse directories.',
        dest='recurse',
        action='store_true',
        default=False
    )

    parser.add_argument(
        '-o', '--output',
        help='A path to which the output should be written.',
        dest='output_destination',
        type=FileType('w'),
        default=stdout
    )

    return parser


def main():
    args = get_parser().parse_args()

    print('\n'.join(sorted(extract_metadata_names(sources=args.sources, recurse=args.recurse))))


if __name__ == '__main__':
    main()
