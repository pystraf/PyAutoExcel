import inspect
import types


class Register:
    """
    A simple decorator-based registration system for classes.
    """

    def __init__(self):
        self.engines = {}

    def add(self, engine_name: str, engine):
        """
        Add an engine to the register.

        :param engine_name: The name of the engine.
        :type engine_name: str
        :param engine: The engine class.
        :type engine: type
        """
        self.engines[engine_name] = engine

    def get(self, engine_name: str):
        """
        Get an engine from the register.

        :param engine_name: The name of the engine.
        :type engine_name: str
        :return: The engine class.
        :rtype: type
        :raises KeyError: If the engine is not found.
        """
        if engine_name not in self.engines:
            raise KeyError(f"Engine '{engine_name}' not found.")
        return self.engines[engine_name]

    def add_from_module(
        self, module: types.ModuleType, sub_class_filters: list[type], field_name: str
    ):
        """
        Add engines from a module to the register.

        :param module: The module to search for classes.
        :type module: types.ModuleType
        :param sub_class_filters: A list of base classes to filter the classes by.
        :type sub_class_filters: list[type]
        :param field_name: The name of the field to check for.
        :type field_name: str
        """
        for name in dir(module):
            data = getattr(module, name)
            if inspect.isclass(data):
                if any([issubclass(data, scls) for scls in sub_class_filters]):
                    if hasattr(data, field_name):
                        field = getattr(data, field_name)
                        if field != "":
                            self.add(field, data)
