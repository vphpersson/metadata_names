from pathlib import Path
from zipfile import ZipFile
from typing import Dict, Any

import xmltodict

from metadata_names.extractors import MetadataExtractor, register_metadata_extractor


@register_metadata_extractor
class MSXMLMetadataExtractor(MetadataExtractor):

    MIME_TYPES = {
        'application/vnd.ms-excel.addin.macroEnabled.12',
        'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
        'application/vnd.ms-excel.sheet.macroEnabled.12',
        'application/vnd.ms-excel.template.macroEnabled.12',
        'application/vnd.ms-powerpoint.addin.macroEnabled.12',
        'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
        'application/vnd.ms-powerpoint.slideshow.macroEnabled.12',
        'application/vnd.ms-powerpoint.template.macroEnabled.12',
        'application/vnd.ms-word.document.macroEnabled.12',
        'application/vnd.ms-word.template.macroEnabled.12',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
        'application/vnd.openxmlformats-officedocument.presentationml.template',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
    }

    @staticmethod
    def from_path(path: Path) -> Dict[str, Any]:
        zf = ZipFile(path)

        core_properties = xmltodict.parse(zf.read('docProps/core.xml'))['cp:coreProperties']

        created = core_properties.get('dcterms:created', {})
        modified = core_properties.get('dcterms:modified', {})

        core_metadata = {
            'category': core_properties.get('cp:category'),
            'contentStatus': core_properties.get('cp:contentStatus'),
            'created': created.get('#text', ''),
            'creator': core_properties.get('dc:creator'),
            'description': core_properties.get('dc:description'),
            'identifier': core_properties.get('dc:identifier'),
            'keywords': core_properties.get('cp:keywords'),
            'language': core_properties.get('dc:language'),
            'lastModifiedBy': core_properties.get('cp:lastModifiedBy'),
            'lastPrinted': core_properties.get('cp:lastPrinted'),
            'modified': modified.get('#text', ''),
            'revision': core_properties.get('cp:revision'),
            'subject': core_properties.get('dc:subject'),
            'title': core_properties.get('dc:title'),
            'version': core_properties.get('cp:version'),
        }

        app_properties = xmltodict.parse(zf.read('docProps/app.xml'))['Properties']

        app_metadata = {
            'Company': app_properties.get('Company'),
            'Template': app_properties.get('Template'),
            # There could be more interesting fields.
        }

        # try:
        #     custom_properties = xmltodict.parse(zf.read('docProps/custom.xml'))['Properties']
        #     for custom_property_map in custom_properties['property']:
        #         for key, value in custom_property_map.items():
        #             print(key, value)
        #     # c.update(custom_properties.keys())
        # except KeyError:
        #     continue

        return {
            'core': core_metadata,
            'app': app_metadata
        }
