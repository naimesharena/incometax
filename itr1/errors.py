"""Typed input errors that identify the *exact* form field at fault.

Every rupee figure entered by the taxpayer flows through :func:`parse_int` /
:func:`parse_float`.  They are deliberately tolerant of real-world typing:
blank values fall back to a default, ``"12,00,000"`` (Indian or western
grouping), stray spaces, ``"₹ 5000"``, and decimal text such as
``"12000.50"`` are all accepted.  Anything genuinely unparseable raises
:class:`DataError` whose message names the offending field, for example::

    Standard deduction override (Salary income): 'abc' is not a valid
    whole number — enter digits only (e.g. 75000 or 75,000).

The current "field scope" (section of the form being processed) is tracked in
a thread-local stack via :class:`field_scope`, so deeply nested parsing
helpers always know which human-visible field they are working on.
"""

from __future__ import annotations

import re
import threading
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

__all__ = ["DataError", "field_scope", "parse_int", "parse_float", "resolve_field"]

_state = threading.local()


class DataError(ValueError):
    """A taxpayer-supplied value could not be understood.

    Attributes
    ----------
    field : str
        Human label of the field, e.g. ``"Standard deduction override
        (Salary income)"``.
    path : str
        Best-effort wizard model path, e.g.
        ``"salary.standard_deduction_override"`` — lets the UI highlight the
        exact input.  May be empty.
    value :
        The offending raw value.
    """

    def __init__(self, message: str, field: str = "", path: str = "", value=None):
        super().__init__(message)
        self.field = field
        self.path = path
        self.value = value

    def __str__(self):
        base = super().__str__()
        return f"{self.field}: {base}" if self.field else base


class field_scope:
    """Context manager pushing ``(label, path_prefix)`` onto the field stack.

    Usage::

        with field_scope("Salary income", "salary"):
            ...  # any parse error here is attributed to the salary section
    """

    __slots__ = ("label", "path")

    def __init__(self, label: str, path: str = ""):
        self.label = label
        self.path = path

    def __enter__(self):
        stack = getattr(_state, "stack", None)
        if stack is None:
            stack = _state.stack = []
        stack.append((self.label, self.path))
        return self

    def __exit__(self, exc_type, exc, tb):
        _state.stack.pop()
        return False


def _stack():
    return getattr(_state, "stack", None) or []


def reset_scopes():
    """Clear the field-scope stack (call at the start of a build request so a
    previous request aborted mid-scope on the same thread cannot leak labels)."""
    _state.stack = []


def resolve_field(key_label: str | None = None, key_path: str | None = None,
                  field: str | None = None, path: str | None = None):
    """Combine the live scope stack with per-key info into (label, path)."""
    stack = _stack()
    labels = [lbl for lbl, _ in stack]
    if field:
        labels.append(field)
    elif key_label:
        labels.append(key_label)
    label = labels[-1] if labels else (field or key_label or "Input")
    if len(labels) > 1:
        label = f"{label} ({' › '.join(labels[:-1])})"

    # machine path: outermost non-empty prefix wins; append key leaf
    path_parts = ""
    for _, p in stack:
        if p:
            path_parts = f"{path_parts}.{p}" if path_parts else p
    leaf = path or key_path or ""
    if leaf:
        path_parts = f"{path_parts}.{leaf}" if path_parts else leaf
    return label, path_parts


_NUM_RE = re.compile(r"-?\d+(\.\d+)?")
_GROUPED_RE = re.compile(r"-?\d{1,3}(,\d{2,3})+(\.\d+)?")


def _normalise(value):
    """Return a clean numeric string, or None when blank, or raise KeyError-ish."""
    s = str(value).strip()
    s = (s.replace("₹", "").replace("Rs.", "").replace("Rs", "")
          .replace("INR", "").strip())
    if not s:
        return None
    if _GROUPED_RE.fullmatch(s):
        return s.replace(",", "")
    if _NUM_RE.fullmatch(s):
        return s
    return _INVALID


class _Invalid:
    pass


_INVALID = _Invalid()


def _to_decimal(value, *, kind, field, path, key_label, key_path, default):
    label, pth = resolve_field(key_label, key_path, field, path)
    if value is None:
        return default
    if isinstance(value, bool):
        return Decimal(int(value))
    if isinstance(value, (int, float, Decimal)):
        try:
            d = Decimal(str(value))
        except InvalidOperation:  # pragma: no cover - defensive
            raise DataError(f"'{value}' is not a valid number.",
                            field=label, path=pth, value=value) from None
        if not d.is_finite():  # NaN / inf
            raise DataError(f"'{value}' is not a valid number.",
                            field=label, path=pth, value=value)
        return d
    s = _normalise(value)
    if s is None:
        return default
    if s is _INVALID:
        raw = str(value).strip()
        want = "a whole number (rupees)" if kind == "int" else "a number"
        raise DataError(
            f"'{raw}' is not {want} — enter digits only "
            f"(commas are fine, e.g. 75000 or 75,000; remove any letters/symbols).",
            field=label, path=pth, value=value)
    return Decimal(s)


def parse_int(value, *, field=None, path=None, key_label=None, key_path=None,
              default=0):
    """Tolerant whole-rupee parser.

    Blank → ``default``; ``"12,00,000"`` → 1200000; ``"12000.50"`` → 12001
    (rounded half-up); anything else → :class:`DataError` naming the field.
    """
    dec = _to_decimal(value, kind="int", field=field, path=path,
                      key_label=key_label, key_path=key_path, default=None)
    if dec is None:
        return default
    return int(dec.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def parse_float(value, *, field=None, path=None, key_label=None, key_path=None,
                default=0.0):
    """Tolerant float parser (percentages such as co-owner share)."""
    dec = _to_decimal(value, kind="float", field=field, path=path,
                      key_label=key_label, key_path=key_path, default=None)
    if dec is None:
        return default
    return float(dec)
