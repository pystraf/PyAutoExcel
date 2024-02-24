import re
from abc import ABC, abstractmethod
from typing import Any, Callable, Union


class BaseValidator(ABC):
    """
    Base class for validators.

    Subclasses should implement the `validate` method.
    """

    def __init__(self):
        pass

    @abstractmethod
    def validate(self, data):
        """
        Validate the given data.

        :param data: The data to validate.
        :raises ValueError: If the data is invalid.
        """
        raise NotImplementedError


class TypeValidator(BaseValidator):
    """
    Validator for data types.

    :param data_types: The allowed data types.
    """

    def __init__(self, *data_types):
        super().__init__()
        self.data_types = data_types

    def validate(self, data):
        if not isinstance(data, self.data_types):
            raise ValueError(
                f"Expected data of type {self.data_types!r}, but got {type(data)}."
            )
        return data


class RangeValidator(BaseValidator):
    """
    Validator for numeric ranges.

    :param min_value: The minimum value of the range.
    :type min_value: int | float
    :param max_value: The maximum value of the range.
    :type max_value: int | float
    """

    def __init__(self, min_value, max_value):
        super().__init__()
        self.min_value = min_value
        self.max_value = max_value

    def validate(self, data):
        if not self.min_value <= data <= self.max_value:
            raise ValueError(
                f"Value {data} is out of range {self.min_value} to {self.max_value}"
            )
        return data


class ContainsValidator(BaseValidator):
    """
    Validator for checking a value contains a specific sequence.

    :param allowed_values: The allowed values.
    :type allowed_values: list | tuple
    """

    def __init__(self, allowed_values: Union[list, tuple]):
        super().__init__()
        self.allowed_values = allowed_values

    def validate(self, data):
        if data not in self.allowed_values:
            raise ValueError(
                f"Value {data} is not allowed, "
                f"allowed values are: {self.allowed_values!r}"
            )
        return data


class CustomValidator(BaseValidator):
    """
    Validator for custom validation logic.

    :param func: The custom validation function.
    :type func: Callable[[Any], bool]
    """

    def __init__(self, func: Callable[[Any], bool]):
        super().__init__()
        self.func = func

    def validate(self, data):
        if not self.func(data):
            raise ValueError(f"Value {data} is invalid")
        return data


class UrlValidator(BaseValidator):
    """
    Validator for checking a value is a valid URL.
    """

    def __init__(self):
        super().__init__()

    def validate(self, data):
        if (
            not re.match(r"^https?://.+$", data)
            and not re.match(r"^http://.+$", data)
            and not re.match(r"^ftp://.+$", data)
            and not re.match(r"^file://.+$", data)
        ):
            raise ValueError(f"Value {data} is not a valid URL")
        return data


class EmailValidator(BaseValidator):
    """
    Validator for checking a value is a valid email address.
    """

    def __init__(self):
        super().__init__()

    def validate(self, data):
        if not re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", data):
            raise ValueError(f"Value {data} is not a valid email address")
        return data


class StringLengthValidator(BaseValidator):
    """
    Validator for string lengths.

    :param min_length: The minimum length of the string.
    :type min_length: int
    :param max_length: The maximum length of the string.
    :type max_length: int
    """

    def __init__(self, min_length: int, max_length: int):
        super().__init__()
        self.min_length = min_length
        self.max_length = max_length

    def validate(self, data):
        if not (self.min_length <= len(data) <= self.max_length):
            raise ValueError(f"Value {data} is not within the allowed length range")
        return data


class MinBoundaryValidator(BaseValidator):
    """
    Validator for checking a value is greater than or equal to a minimum boundary.

    :param min_boundary: The minimum boundary.
    :type min_boundary: Number
    """

    def __init__(self, min_boundary: int):
        super().__init__()
        self.min_boundary = min_boundary

    def validate(self, data):
        if data < self.min_boundary:
            raise ValueError(f"Value {data} is less than the minimum boundary")
        return data


class MaxBoundaryValidator(BaseValidator):
    """
    Validator for checking a value is less than or equal to a maximum boundary.

    :param max_boundary: The maximum boundary.
    :type max_boundary: Number
    """

    def __init__(self, max_boundary: int):
        super().__init__()
        self.max_boundary = max_boundary

    def validate(self, data):
        if data > self.max_boundary:
            raise ValueError(f"Value {data} is greater than the maximum boundary")
        return data


class UniqueValidator(BaseValidator):
    """
    Validator for checking a sequence of values are unique.
    """

    def __init__(self):
        super().__init__()

    def validate(self, data: Union[list, tuple, set, frozenset]):
        if isinstance(data, (set, frozenset)):
            return data
        if len(data) != len(set(data)):
            raise ValueError(f"Value {data} is not unique.")
        return data


class ConcatValidator(BaseValidator):
    """
    Validator for concatenating multiple validators.

    :param validators: The validators to concatenate.
    """

    def __init__(self, *validators: BaseValidator):
        super().__init__()
        self.validators = validators

    def validate(self, data):
        for validator in self.validators:
            data = validator.validate(data)
        return data
