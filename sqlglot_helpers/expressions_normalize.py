"""
Helpers for normalizing sqlglot Expression trees (non-query expressions).

Focus: parentheses-only normalization and a minimal, schema-free simplify path
that is safe for boolean/filter expressions. Also provides utilities to
differentiate full SQL statements (queries) from plain expressions to avoid
running query-oriented optimizer passes on statements.
"""

from __future__ import annotations

from typing import Optional

from sqlglot import expressions as exp


# Classes that indicate a full SQL statement/query at the root.
QUERY_ROOTS = (
    exp.Select,
    exp.SetOperation,  # UNION/INTERSECT/EXCEPT
    exp.With,          # CTEs wrapping a query
    exp.Insert,
    exp.Update,
    exp.Delete,
    exp.Merge,
    exp.Create,
    exp.Drop,
    exp.Alter,
)


def _unwrap_paren(node: exp.Expression) -> exp.Expression:
    """Return the innermost node by removing Paren wrappers."""
    while isinstance(node, exp.Paren):
        node = node.this
    return node


def is_full_query(node: exp.Expression) -> bool:
    """True if the (unwrapped) root is a SQL statement/query node."""
    return isinstance(_unwrap_paren(node), QUERY_ROOTS)


def strip_parens(expr: exp.Expression) -> exp.Expression:
    """
    Remove Paren wrappers from the entire tree without changing structure.

    Returns a new Expression; original is not mutated.
    """
    return expr.copy().transform(lambda n: n.this if isinstance(n, exp.Paren) else n)


def normalize_parentheses_only(
    expr: exp.Expression,
    *,
    dialect: Optional[str] = None,
) -> str:
    """
    Parentheses-only normalization suitable for non-query expressions.

    - Strips exp.Paren nodes
    - Emits normalized, compact SQL (no pretty formatting)

    Raises ValueError if a full statement is provided; adjust if you prefer a no-op.
    """
    if is_full_query(expr):
        raise ValueError("normalize_parentheses_only expects a non-query expression")

    cleaned = strip_parens(expr)
    return cleaned.sql(dialect=dialect, normalize=True, pretty=False)


def normalize_expression(
    expr: exp.Expression,
    *,
    dialect: Optional[str] = None,
    allow_queries: bool = False,
) -> str:
    """
    Normalize a non-query expression with safe, schema-free simplification.

    Steps:
    - Strip redundant parentheses
    - Apply only a minimal simplify pass (if available)
    - Render with normalize=True, pretty=False for stable output

    Set allow_queries=True to bypass the query guard, but note that this helper
    is intended for non-query expressions. It avoids qualification/resolution
    passes that can raise column resolution errors on statements.
    """
    if is_full_query(expr) and not allow_queries:
        raise ValueError("normalize_expression expects a non-query expression")

    cleaned = strip_parens(expr)

    # Apply only the simplify pass when available; avoid full optimize pipelines.
    try:
        # Prefer calling the pass function directly to avoid optimizer registry differences.
        from sqlglot.optimizer import simplify as _simplify

        simplified = _simplify.simplify(cleaned)
    except Exception:
        # Fallback to just paren-stripped when optimizer API differs or is unavailable.
        simplified = cleaned

    return simplified.sql(dialect=dialect, normalize=True, pretty=False)


def paren_insensitive_equal(a: exp.Expression, b: exp.Expression) -> bool:
    """AST equality that ignores only parentheses wrappers (exp.Paren)."""
    return strip_parens(a) == strip_parens(b)


__all__ = [
    "QUERY_ROOTS",
    "is_full_query",
    "strip_parens",
    "normalize_parentheses_only",
    "normalize_expression",
    "paren_insensitive_equal",
]

