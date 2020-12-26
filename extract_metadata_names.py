#!/usr/bin/env python3

from pathlib import Path
from argparse import ArgumentParser, FileType
from sys import stdout
from logging import DEBUG

from terminal_utils.progressor import Progressor
from terminal_utils.log_handlers import ColoredProgressorLogHandler

from metadata_names import extract_metadata_names, LOG
from metadata_names.errors import UnsupportedMimetypeError


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


def main():
    args = MetadataNamesArgumentParser().parse_args()

    try:
        with Progressor() as progressor:
            if args.quiet:
                LOG.disabled = True
            else:
                handler = ColoredProgressorLogHandler(progressor=progressor, print_warnings=not args.ignore_warnings)
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
