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
    sort_in_items: bool = True,
) -> str:
    """
    Normalize a non-query expression with safe, schema-free simplification and
    canonical ordering for commutative boolean groups.

    Steps:
    - Strip redundant parentheses
    - Apply only a minimal simplify pass (if available)
    - Canonicalize ordering within AND/OR groups (and optionally IN lists)
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

    # Canonicalize commutative boolean groups and (optionally) IN list items.
    canon = _canonicalize_boolean_groups(simplified, dialect=dialect, sort_in_items=sort_in_items)

    return canon.sql(dialect=dialect, normalize=True, pretty=False)


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


# ---- Internal canonicalization helpers (ordering for commutative groups) ----

def _stable_key(e: exp.Expression, dialect: Optional[str]) -> str:
    """
    Stable sort key for child expressions: strip redundant parentheses and render
    with normalized casing/quoting so ordering is deterministic.
    """
    return strip_parens(e).sql(dialect=dialect, normalize=True, pretty=False)


def _flatten_boolean(node: exp.Expression, kind: type[exp.Expression]) -> list[exp.Expression]:
    """
    Collect all child terms under a chain of the same boolean operator (AND/OR).
    Returns the list of leaf expressions in no particular order.
    """
    parts: list[exp.Expression] = []

    def collect(n: exp.Expression):
        if isinstance(n, kind):
            collect(n.this)
            collect(n.expression)
        else:
            parts.append(n)

    collect(node)
    return parts


def _rebuild_boolean(kind: type[exp.Expression], parts: list[exp.Expression]) -> exp.Expression:
    """Rebuild a boolean chain (AND/OR) from a list of parts."""
    if not parts:
        raise ValueError("Cannot rebuild boolean node from empty parts")
    if len(parts) == 1:
        return parts[0]
    return exp.and_(*parts) if kind is exp.And else exp.or_(*parts)


def _canonicalize_boolean_groups(
    root: exp.Expression,
    *,
    dialect: Optional[str],
    sort_in_items: bool,
) -> exp.Expression:
    """
    Make ordering of commutative boolean groups deterministic:
      - strip redundant parentheses (caller already did)
      - sort terms in AND and OR groups
      - optionally sort IN(...) list items
    """
    def visit(n: exp.Expression) -> exp.Expression:
        # First normalize children bottom-up
        n = n.transform(visit)

        if isinstance(n, exp.And):
            parts = _flatten_boolean(n, exp.And)
            parts.sort(key=lambda x: _stable_key(x, dialect))
            return _rebuild_boolean(exp.And, parts)

        if isinstance(n, exp.Or):
            parts = _flatten_boolean(n, exp.Or)
            parts.sort(key=lambda x: _stable_key(x, dialect))
            return _rebuild_boolean(exp.Or, parts)

        # Optional: sort items inside IN(...) for stability.
        if sort_in_items and isinstance(n, exp.In) and n.args.get("expressions"):
            items = list(n.args["expressions"])
            items.sort(key=lambda x: _stable_key(x, dialect))
            return exp.In(this=n.this, expressions=items, not_=n.args.get("not"))

        return n

    return visit(root)
