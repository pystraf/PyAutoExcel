from typing import Any

from .Validators import BaseValidator


class Field:
    """
    Represents a field in a form.

    :param name: The name of the field.
    :param validator: The validator for the field.
    :param required: Whether the field is required.
    :param default: The default value for the field.
    """

    validator: BaseValidator
    name: str
    value: Any

    def __init__(
        self,
        name: str,
        validator: BaseValidator,
        required: bool = False,
        default: Any = None,
    ):
        if required and default:
            raise ValueError("A required field cannot have a default value.")

        self.required = required
        self.validator = validator
        self.name = name
        self.value = default
        self.default = default

    def set(self, value):
        """
        Sets the value of the field.

        :param value: The value to set.
        :raises ValidationError: If the value is invalid.
        """
        self.validator.validate(value)
        self.value = value

    def get(self):
        """
        Gets the value of the field.

        :return: The value of the field.
        """
        return self.value
