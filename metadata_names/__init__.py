from pathlib import Path
from typing import Iterable
from logging import getLogger
from unicodedata import category as unicodedata_category

from magic import Magic
from terminal_utils.log_handlers import ProgressStatus
from metadata_extractor.extractors.office_open_xml import OfficeOpenXMLMetadataExtractor
from metadata_extractor.extractors.ole import OLEMetadataExtractor
from metadata_extractor.extractors.pdf import PDFMetadataExtractor
from metadata_extractor.extractors import MetadataExtractor

from metadata_names.errors import UnsupportedMimetypeError

LOG = getLogger(__name__)


def extract_metadata_names(sources: Iterable[Path], recurse: bool, exit_on_unsupported_mimetype: bool = False) -> set[str]:
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

    names: set[str] = set()

    try:
        for i, path in enumerate(paths, start=1):
            LOG.debug(ProgressStatus(iteration=i-1, total=len(paths), prefix='Extracting metadata... '))

            try:
                mimetype = get_mimetype(str(path))
            except FileNotFoundError as e:
                LOG.warning(e)
                continue

            try:
                metadata_extractor = MetadataExtractor.extractor_from_mime_type(mime_type=mimetype)
            except KeyError:
                error_message = f'Unsupported mimetype {mimetype} for file {path}'
                if exit_on_unsupported_mimetype:
                    raise UnsupportedMimetypeError(observed_mimetype=mimetype, file_path=str(path))
                else:
                    LOG.warning(error_message)
                    continue

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
    except KeyboardInterrupt:
        pass

    return {
        ''.join(character for character in name if unicodedata_category(character) in {'Lu', 'Ll', 'Zs'}).strip()
        for name in names
    }
