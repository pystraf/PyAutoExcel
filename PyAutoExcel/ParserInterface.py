"""
Fill members into an instance of the class.

Created at: 2023/01/25 12:01:00
Written by: pystraf

Note: This module deprecated since 3.0.3, Will be removed in 3.1.0
"""
import warnings
from typing import Union

from . import Utils

warnings.warn(
    DeprecationWarning(
        "This module deprecated since 3.0.3, "
        "Will be removed in 3.1.0. "
        "Use ParserInterface2 instead."
    )
)


class Field(metaclass=Utils.FinalMeta):
    """
    Field class
    field_name -> field name (unique)
    field_type -> field type (e.g 'int' or '(int, float)')

    :param field_name: The name of the field
    :type field_name: str
    :param field_type: The type or tuple-type of the field
    :type field_type: Union[type, tuple[type, ...]]
    """

    def __init__(self, field_name: str, field_type: Union[type, tuple[type, ...]]):
        """
        Init field object
        :param field_name: The name of the field
        :param field_type: The type or tuple-type of the field
        """
        self.name, self.type = field_name, field_type

    def __repr__(self):
        params = f"(field_name={self.name!r}, field_type={self.type!r})"
        return f"PyAutoExcel.ParseInterface.Field{params}"


class Parser(metaclass=Utils.FinalMeta):
    """
    fill members to object

    members:
    fields -> list of Field object
    records -> [[fieldname, fieldtype, fieldvalue], ...]
    result -> parsed result object

    :param fields: list of Field object.
    :type fields: list[Field]
    """

    def __init__(self, fields: list[Field]):
        """
        Init members parser
        :param fields: list of Field object
        """
        self.fields = fields.copy()
        self.records = []
        self.result = None
        self.errors = [
            "The length of the parameter (%d) must be equal to the length of the field (%d).",
            "Excepted %s, got %s.",
        ]

    def prepare(self):
        """
        Preparing basic field info
        It's fill 'records' likes: [[fieldname, fieldtype, None], ...]

        :return: Nothing
        """
        for f in self.fields:
            self.records.append([f.name, f.type, None])

    def fill(self, params: list):
        """
        fill to params to records.
        :param params: list of params.
        :type params: list
        :raises AssertionError: param's length is not equal to field's length.

        :return: Nothing
        """
        assert len(params) == len(self.fields), self.errors[0] % (
            len(params),
            len(self.fields),
        )
        for i, value in enumerate(params):
            # records[i][0] -> fieldname
            # records[i][1] -> fieldtype
            # records[i][2] -> fieldvalue
            self.records[i][2] = value

    def verify(self):
        """
        Verify fields is specific type.

        :raises AssertionError: if type of params is not match with type of field.
        :return: Nothing
        """
        for r in self.records:
            cls = Utils.get_cls_name(objs=r[2])
            error_msg = self.errors[1] % (
                Utils.humanize_items(items=cls, using_or=True),
                Utils.object_class_name(obj=r[1]),
            )
            assert isinstance(r[2], r[1]), error_msg

    def parse(self, target_cls: type):
        """
        Fill data to specific class and return an object of this class.

        :param target_cls: The class want to be filled.
        :type target_cls: type
        :return: an object of target class.
        """
        self.result = target_cls()
        for r in self.records:
            setattr(self.result, r[0], r[2])

    def __repr__(self):
        return f"PyAutoExcel.ParseInterface.Parser(fields={self.fields!r})"


def parse(fields: list[Field], params: list, target_cls: type):
    """
    Fill fields to target class and return an object.

    :param fields: list of Field object
    :type fields: list[Field]
    :param params: list of params.
    :type params: list
    :param target_cls: target class.
    :type target_cls: type
    :return: an object of target class

    :raise AssertionError: if type of params is not match with type of field.
    """
    # order:
    # 1. create parser
    # 2. preparing basic data
    # 3. fill values
    # 4. verify type.
    # 5. parse to target
    parser = Parser(fields=fields)
    parser.prepare()
    parser.fill(params=params)
    parser.verify()
    parser.parse(target_cls=target_cls)
    return parser.result
