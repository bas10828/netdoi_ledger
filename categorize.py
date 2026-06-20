"""Category guessing for slip memo text. Categories + keywords are stored in the `categories`
table (admin-editable via /settings/categories), not hardcoded here.
"""

import psycopg2.extras


def get_categories(cur):
    """Works regardless of the calling cursor's factory — opens its own RealDictCursor."""
    with cur.connection.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as c:
        c.execute("SELECT name, keywords FROM categories ORDER BY sort_order, id")
        return [{"name": r["name"], "keywords": list(r["keywords"] or [])} for r in c.fetchall()]


def guess_category(memo, categories):
    if not memo:
        return None
    text = memo.lower()
    for cat in categories:
        if any(kw.lower() in text for kw in cat["keywords"]):
            return cat["name"]
    return None
