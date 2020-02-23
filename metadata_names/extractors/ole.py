from pathlib import Path
from typing import Dict, Any

import olefile

from metadata_names.extractors import MetadataExtractor, register_metadata_extractor


@register_metadata_extractor
class OLEMetadataExtractor(MetadataExtractor):

    MIME_TYPES = {
        'application/msword',
        'application/vnd.ms-excel',
        'application/vnd.ms-powerpoint',
    }

    @staticmethod
    def from_path(path: Path) -> Dict[str, Any]:
        ole = olefile.OleFileIO(path)

        summary_metadata = {}
        if ole.exists('\x05SummaryInformation'):
            props = ole.getproperties('\x05SummaryInformation')
            summary_metadata = {
                attribute: props.get(i + 1, None)
                for i, attribute in enumerate(olefile.OleMetadata.SUMMARY_ATTRIBS)
            }

        document_summary_metadata = {}
        if ole.exists('\x05DocumentSummaryInformation'):
            props = ole.getproperties('\x05DocumentSummaryInformation')
            document_summary_metadata = {
                attribute: props.get(i + 1, None)
                for i, attribute in enumerate(olefile.OleMetadata.DOCSUM_ATTRIBS)
            }

        return {
            'SummaryInformation': (
                summary_metadata or {
                    attribute: None
                    for attribute in olefile.OleMetadata.SUMMARY_ATTRIBS
                }
            ),
            'DocumentSummaryInformation': (
                document_summary_metadata or {
                    attribute: None
                    for attribute in olefile.OleMetadata.DOCSUM_ATTRIBS
                }
            )
        }
