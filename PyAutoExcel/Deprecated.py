"""
Deprecated class for information on abandoned functions.
"""
import functools
import types

from colorama import Fore

from .Utils import init_colorama

init_colorama()

WARNINFO = "engine '%s' is deprecated and will be removed in %s, use %s instead."


class DeprecatedInfo:
    """
    A class to represent deprecated information for an engine.

    :param is_deprecated: Indicates if the engine is deprecated.
    :type is_deprecated: bool
    :param remove_version: The version in which the engine will be removed.
    :type remove_version: str
    :param instead: Suggested alternative for the deprecated engine.
    :type instead: str
    """

    def __init__(
        self,
        is_deprecated: bool = False,
        remove_version: str = "",
        instead: str = "",
    ):
        self.is_deprecated = is_deprecated
        self.remove_version = remove_version
        self.instead = instead

    def show(self, engine: str):
        """
        Prints a deprecation warning message for the specified engine.

        :param engine: The name of the deprecated engine.
        :type engine: str
        """
        msg = WARNINFO % (engine, self.remove_version, self.instead)
        print("%sWarning: %s" % (Fore.YELLOW, msg))


class DeprecationMessages:
    """
    A collection of messages related to deprecation.
    """

    @staticmethod
    def deprecated(name: str, version: str, remove_version: str) -> str:
        """
        Generate a deprecation message for a named element.

        :param name: The name of the deprecated element.
        :param version: The version in which the element was deprecated.
        :param remove_version: The version in which the element will be removed.
        :return: The deprecation message.
        """
        return f"{name} is deprecated since {version} and will be remove in {remove_version}"

    @staticmethod
    def deprecated_with_replacement(
        name: str, version: str, remove_version: str, replacement: str
    ):
        """
        Generate a deprecation message for a named element with a replacement.

        :param name: The name of the deprecated element.
        :param version: The version in which the element was deprecated.
        :param remove_version: The version in which the element will be removed.
        :param replacement: The replacement for the deprecated element.
        :return: The deprecation message with replacement information.
        """
        origin = DeprecationMessages.deprecated(name, version, remove_version)
        return f"{origin}, use {replacement} instead."


def func_deprecated(message: str, warn_id: int):
    """
    Decorator for marking a function as deprecated.

    :param message: The deprecation message.
    :param warn_id: The warning ID.
    :return: The decorated function.
    """

    def inner(func: types.FunctionType):
        @functools.wraps(func)
        def wrapped(*args, **kwargs):
            if not hasattr(func, "__refcount__"):
                func.__refcount__ = 1
            else:
                func.__refcount__ += 1
            if func.__refcount__ == 1:
                print(f"{Fore.YELLOW}Warning ({hex(warn_id)}): {message}")
            return func(*args, **kwargs)

        return wrapped

    return inner


class_deprecated_refcount = {}
class_deprecated_message = {}


class DeprecatedMeta(type):
    """
    Metaclass for handling deprecated classes.
    """

    def __call__(cls, *args, **kwargs):
        global class_deprecated_message, class_deprecated_refcount
        if cls not in class_deprecated_refcount:
            class_deprecated_refcount[cls] = 1
        else:
            class_deprecated_refcount[cls] += 1
        if class_deprecated_refcount[cls] == 1:
            print(
                f"{Fore.YELLOW}Warning "
                f"({hex(class_deprecated_message[cls][1])}): "
                f"{class_deprecated_message[cls][0]}"
            )
        obj = type.__call__(cls, *args, **kwargs)
        return obj


def set_deprecated_msg(cls: DeprecatedMeta, msg: str, warn_id: int):
    """
    Set the deprecated message for a class.

    :param warn_id: The type ID of the warning.
    :param cls: The class that is being deprecated.
    :param msg: The message to display when the class is used.
    """
    global class_deprecated_message
    class_deprecated_message[cls] = (msg, warn_id)


def process_deprecated(info: DeprecatedInfo, name: str):
    """
    Checks if the engine is deprecated and displays a warning if so.

    :param info: A DeprecatedInfo object containing deprecation information about the engine.
    :type info: DeprecatedInfo
    :param name: The name of the engine being checked.
    :type name: str
    """
    if info.is_deprecated:
        info.show(engine=name)
