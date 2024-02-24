from .Field import Field


class Parser:
    """
    Parser class.

    :param params: dict of parameters.
    :type params: dict
    :param target: class of the object to be parsed.
    :type target: type
    :param fields: list of fields to be parsed.
    :type fields: list[Field]
    """

    def __init__(self, params: dict, target: type, fields: list[Field]):
        self.fields = fields
        self.target = target
        self.params = params

    def parse(self, *args, **kwargs):
        """
        Parse the parameters and return an instance of the target class.

        :return: instance of the target class.
        """
        inst = self.target(*args, **kwargs)
        for field in self.fields:
            if field.name in self.params:
                field.validator.validate(self.params[field.name])
                setattr(inst, field.name, self.params[field.name])
            elif field.required:
                raise ValueError(f"Required field {field.name} not provided.")
            else:
                setattr(inst, field.name, field.default)
        return inst
