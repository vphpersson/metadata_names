from __future__ import annotations
from abc import ABC, abstractmethod
from typing import ClassVar, Set, Type, Dict, Any
from pathlib import Path


class MetadataExtractor(ABC):
    MIME_TYPES: ClassVar[Set[str]] = NotImplemented
    MIME_TYPE_TO_EXTRACTOR_CLASS: Dict[str, Type[MetadataExtractor]] = {}

    @staticmethod
    @abstractmethod
    def from_path(path: Path) -> Dict[str, Any]:
        raise NotImplementedError

    @classmethod
    def extractor_from_mime_type(cls, mime_type: str) -> Type[MetadataExtractor]:
        return cls.MIME_TYPE_TO_EXTRACTOR_CLASS[mime_type]


def register_metadata_extractor(cls: Type[MetadataExtractor]) -> Type[MetadataExtractor]:
    cls.MIME_TYPE_TO_EXTRACTOR_CLASS.update({mime_type: cls for mime_type in cls.MIME_TYPES})
    return cls
