from pathlib import Path
from typing import Dict

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument

from metadata_names.extractors import MetadataExtractor, register_metadata_extractor


@register_metadata_extractor
class PDFMetadataExtractor(MetadataExtractor):
    MIME_TYPES = {'application/pdf'}

    # TODO: Add proper type hints.
    @staticmethod
    def from_path(path: Path) -> Dict[str, bytes]:
        with path.open('rb') as fp:
            return next(iter(PDFDocument(PDFParser(fp)).info), dict())
