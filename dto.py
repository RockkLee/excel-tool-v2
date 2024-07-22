from dataclasses import dataclass
from abc import ABC, abstractmethod  # ABC: Abstract Base Class


@dataclass
class ColumnDto(ABC):
    ls: list[str]

    def __new__(cls, *args, **kwargs):
        if cls == ColumnDto:  # or cls.__bases__[0] == ColumnDto:
            raise TypeError("Cannot instantiate abstract class.")
        return super().__new__(cls)


@dataclass
class SupplierNameCol(ColumnDto):
    ls: list[str]
    __name__ = 'Supplier Name'


@dataclass
class RawMaterialCol(ColumnDto):
    ls: list[str]
    __name__ = 'Raw Material #'


@dataclass
class ColorCol(ColumnDto):
    ls: list[str]
    __name__ = 'Color'


@dataclass
class StyleCol(ColumnDto):
    ls: list[str]
    __name__ = 'Style'


@dataclass
class ComponentCol(ColumnDto):
    ls: list[str]
    __name__ = 'Component'
