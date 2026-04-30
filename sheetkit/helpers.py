from typing import Any
from pathlib import Path
from os import PathLike
import inspect

from .const import WARNINGS

# from alx import *

def reprint(line: str, par: Any = None) -> None:
    """
    Rewrites a line on the console without breaking the line. Can be used when you need to rewrite part of a string,
    displayed on the screen. For example, update the completion percentage on the same line:
        import time
        txt = 'Percentage complete: {}%'
        print()
        reprint(txt, 1)
        time.sleep(1)
        reprint(txt, 50)
        time.sleep(1)
        reprint(txt, 100)
    Args:
        line: Line pattern
        par: Parameter
    Return: None
    """
    print('\r', end='')
    if par is None:
        print(line, end='', flush=True)
    else:
        print(line.format(par), end='', flush=True)


class PercentProgress:
    """
    When calling update(), the object prints a string with the progress percentage to the console
    and updates the percentage.
    """
    def __init__(self, total: int, *, times: int = 100, msg: str = '{:.0f}%') -> None:
        """
        Sets parameters
        Parameters
        ----------
        total - total expected iterations
        times - total counter progress updates (e.g., 100)
        msg - message like 'Completed {:.0f}%' or 'Completed {:.1f}%'.
        Accordingly, the output will be: 23% or 23.1% completed
        """
        self.times = times
        self.total = total
        self.msg = msg

        self.mark_list = []

        if int(total / times) >= 1:
            self.mark_list += list(range(0, total, int(total / times)))
        self.mark_list += [total - 1]

    def __str__(self) -> str:
        return f'PercentProgress(msg={self.msg}, times={self.times}, total={self.total})'

    def update(self, counter: int, total: int = 0) -> None:

        if total > 0:
            self.total = total

        if counter == -1:  # для первого появления строки, если нужно вставить перед этим пустую строку
            print()
            reprint(self.msg.format(0) + '...')

        elif counter == 0:
            reprint(self.msg.format(0) + '   ')

        elif counter in self.mark_list and (self.total - 1) > 0:
            percent = (counter / (self.total - 1)) * 100
            reprint(self.msg.format(percent))


def inspect_name() -> str:
    """
    Returns the name of the function from which inspect_name() was called.
    Returns
    -------
    """
    frame = inspect.currentframe()
    if frame is None or frame.f_back is None:
        return "<unknown>()"
    return frame.f_back.f_code.co_name + '()'


def inspect_upper_name() -> str:
    """
    Returns the name of the function that called the function that called inspect_name().

    Returns
    -------
    function name (str)
    """
    frame = inspect.currentframe()
    if frame is None or frame.f_back is None or frame.f_back.f_back is None:
        return "<unknown>()"
    return frame.f_back.f_back.f_code.co_name + '()'  # caller of caller


def wrn(*msg: Any) -> None:
    """Print warning when warnings are enabled."""
    if WARNINGS:
        print("#Warning:", *msg)


def dbg(*string: Any, loc: bool = True, mark: str = "") -> None:
    """Print debug message with optional location marker."""
    loc_str = inspect_upper_name() if loc else ""
    mark_str = f"[{mark}] " if mark else ""
    fil_str = " -> " if (mark or loc) else ""
    print(f"#dbg: {mark_str}{loc_str}{fil_str}", *string)


def add_ext(path: str | PathLike, ext: str) -> str:
    """
    Adds a filename extension if the current extension does not match the specified one.

    Parameters
    ----------
    path - filename with extension, with or without path
    ext - file extension

    Returns
    -------
    """
    # cast to Path
    p = Path(path)

    # ensure that the extension starts with "."
    if not ext.startswith("."):
        ext = "." + ext

    # current file extension (e.g. ".txt")
    current_ext = p.suffix

    if current_ext.lower() == ext.lower():
        # if it matches, simply replace the extension with the "canonical" ext
        new_path = p.with_suffix(ext)
    else:
        # if different, add the new extension to the existing one
        new_path = p.with_name(p.name + ext)

    return str(new_path)
