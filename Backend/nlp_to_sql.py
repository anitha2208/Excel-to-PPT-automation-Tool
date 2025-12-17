# core/nlp_to_sql.py
import re
import pandas as pd

def normalize_column_name(col):
    return re.sub(r'[^a-z0-9]', '', col.lower())

def normalize_text(text):
    return re.sub(r'[^a-z0-9\s]', '', text.lower()).strip()

# ---------------- Condition Parsing ----------------
def parse_conditions(condition_text, norm_map):
    conditions = []
    operators = []
    parts = re.split(r'\b(and|or)\b', condition_text, flags=re.IGNORECASE)

    for i, part in enumerate(parts):
        part = part.strip()
        if part.lower() in ['and', 'or']:
            operators.append(part.upper())
        elif part:
            cond = parse_single_condition(part, norm_map)
            if cond:
                conditions.append(cond)
    return conditions, operators

def parse_single_condition(condition_text, norm_map):
    condition_text = condition_text.strip()

    # BETWEEN
    between_pattern = r'(\w+)\s+(?:is\s+)?between\s+(\d+(?:\.\d+)?)\s+(?:to|and)\s+(\d+(?:\.\d+)?)'
    between_match = re.search(between_pattern, condition_text, re.IGNORECASE)
    if between_match:
        col_name = between_match.group(1).lower()
        min_val = float(between_match.group(2))
        max_val = float(between_match.group(3))
        norm_col = normalize_column_name(col_name)
        if norm_col in norm_map:
            return (norm_map[norm_col], 'BETWEEN', (min_val, max_val))

    # LIKE (starts/ends/contains)
    like_patterns = [
        (r'(\w+)\s+starts?\s+(?:with\s+)?([a-zA-Z0-9]+)', 'STARTS_WITH'),
        (r'(\w+)\s+ends?\s+(?:with\s+)?([a-zA-Z0-9]+)', 'ENDS_WITH'),
        (r'(\w+)\s+contains?\s+([a-zA-Z0-9]+)', 'CONTAINS'),
    ]
    for pattern, op_type in like_patterns:
        match = re.search(pattern, condition_text, re.IGNORECASE)
        if match:
            col_name = match.group(1).lower()
            value = match.group(2)
            norm_col = normalize_column_name(col_name)
            if norm_col in norm_map:
                return (norm_map[norm_col], op_type, value)

    # Comparisons
    comparison_patterns = [
        (r'(\w+)\s+(?:is\s+)?(?:greater\s+than|>)\s+(\d+(?:\.\d+)?)', '>'),
        (r'(\w+)\s+(?:is\s+)?(?:less\s+than|<)\s+(\d+(?:\.\d+)?)', '<'),
        (r'(\w+)\s+(?:is\s+)?(?:greater\s+than\s+or\s+equal\s+to|>=)\s+(\d+(?:\.\d+)?)', '>='),
        (r'(\w+)\s+(?:is\s+)?(?:less\s+than\s+or\s+equal\s+to|<=)\s+(\d+(?:\.\d+)?)', '<='),
        (r'(\w+)\s+(?:is\s+)?(?:equal\s+to|equals?|=)\s+(\d+(?:\.\d+)?)', '='),
        (r'(\w+)\s+(?:is\s+)?(?:not\s+equal\s+to|!=|<>)\s+(\d+(?:\.\d+)?)', '!='),
    ]
    for pattern, operator in comparison_patterns:
        match = re.search(pattern, condition_text, re.IGNORECASE)
        if match:
            col_name = match.group(1).lower()
            value = float(match.group(2))
            norm_col = normalize_column_name(col_name)
            if norm_col in norm_map:
                return (norm_map[norm_col], operator, value)

    # IS NOT text
    not_patterns = [
        r'(\w+)\s+is\s+not\s+([\'"]?)([^\'"\s]+)\2',
        r'(\w+)\s+(?:is\s+)?not\s+equal\s+(?:to\s+)?([\'"]?)([^\'"\s]+)\2',
    ]
    for pattern in not_patterns:
        match = re.search(pattern, condition_text, re.IGNORECASE)
        if match:
            col_name = match.group(1).lower()
            value = match.group(3) if len(match.groups()) >= 3 else match.group(2)
            norm_col = normalize_column_name(col_name)
            if norm_col in norm_map:
                return (norm_map[norm_col], '!=', value)

    # Equality text
    equality_patterns = [
        r'(\w+)\s+(?:is|equals?)\s+([\'"]?)([^\'"\s]+)\2',
        r'(\w+)\s*=\s*([\'"]?)([^\'"\s]+)\2',
        r'(?:whose|where)\s+(\w+)\s+(?:is|equals?)\s+([\'"]?)([^\'"\s]+)\2',
    ]
    for pattern in equality_patterns:
        match = re.search(pattern, condition_text, re.IGNORECASE)
        if match:
            col_name = match.group(1).lower()
            value = match.group(3) if len(match.groups()) >= 3 else match.group(2)
            norm_col = normalize_column_name(col_name)
            if norm_col in norm_map:
                return (norm_map[norm_col], '=', value)

    return None

# ---------------- Aggregation, Order, Limit ----------------
def parse_aggregation(query_text, norm_map):
    query_lower = query_text.lower()
    agg_patterns = [
        (r'sum\s+of\s+(\w+)', 'sum'),
        (r'average\s+of\s+(\w+)', 'mean'),
        (r'avg\s+of\s+(\w+)', 'mean'),
        (r'min(?:imum)?\s+(\w+)', 'min'),
        (r'max(?:imum)?\s+(\w+)', 'max'),
        (r'count\s+of\s+(\w+)', 'count')
    ]
    for pattern, agg_func in agg_patterns:
        match = re.search(pattern, query_lower)
        if match:
            col_name = match.group(1).lower()
            norm_col = normalize_column_name(col_name)
            if norm_col in norm_map:
                return [(norm_map[norm_col], agg_func)]
    return None

def parse_order_by(query_text, norm_map):
    query_lower = query_text.lower()
    order_patterns = [
        (r'(?:increasing|ascending)\s+(?:order\s+of|by)\s+(\w+)', 'ASC'),
        (r'(?:decreasing|descending)\s+(?:order\s+of|by)\s+(\w+)', 'DESC'),
    ]
    for pattern, order_dir in order_patterns:
        match = re.search(pattern, query_lower)
        if match:
            col_name = match.group(1).lower()
            norm_col = normalize_column_name(col_name)
            if norm_col in norm_map:
                return (norm_map[norm_col], order_dir)
    return None

def parse_limit(query_text):
    limit_patterns = [
        r'(?:first|top)\s+(\d+)',
        r'limit\s+(\d+)',
        r'(\d+)\s+(?:rows?|records?)',
    ]
    for pattern in limit_patterns:
        match = re.search(pattern, query_text.lower())
        if match:
            return int(match.group(1))
    return None

# ---------------- SQL Generation ----------------
def generate_sql(user_query: str, df):
    df_cols = list(df.columns)
    norm_map = {normalize_column_name(c): c for c in df_cols}

    user_query_lower = user_query.lower()
    selected_cols = []
    conditions, operators = [], []
    is_distinct = False
    is_count = False
    count_column = None
    aggregations = parse_aggregation(user_query, norm_map)
    order_by = parse_order_by(user_query, norm_map)
    limit = parse_limit(user_query)

    if any(keyword in user_query_lower for keyword in ['distinct', 'unique']):
        is_distinct = True

    if "count" in user_query_lower and not aggregations:
        is_count = True

    if not aggregations and not is_count:
        for norm, orig in norm_map.items():
            if norm in user_query_lower:
                selected_cols.append(orig)
        if not selected_cols:
            selected_cols = df_cols

    where_keywords = ['where', 'whose', 'which', 'that']
    where_part = None
    for keyword in where_keywords:
        if keyword in user_query_lower:
            parts = user_query_lower.split(keyword, 1)
            if len(parts) > 1:
                where_part = parts[1].strip()
                break
    if where_part:
        conditions, operators = parse_conditions(where_part, norm_map)

    sql = "SELECT "
    if aggregations:
        agg_parts = [f"{func.upper()}({col})" for col, func in aggregations]
        sql += ", ".join(agg_parts) + " FROM df"
    elif is_count:
        sql += "COUNT(*) FROM df"
    else:
        sql += (", ".join(selected_cols) if selected_cols else "*") + " FROM df"
        if is_distinct:
            sql = sql.replace("SELECT", "SELECT DISTINCT", 1)

    if conditions:
        cond_str = " AND ".join([f"{col} {op} '{val}'" if op not in ['BETWEEN'] 
                                 else f"{col} BETWEEN {val[0]} AND {val[1]}" 
                                 for col, op, val in conditions])
        sql += f" WHERE {cond_str}"

    if order_by and not aggregations:
        sql += f" ORDER BY {order_by[0]} {order_by[1]}"

    if limit and not aggregations:
        sql += f" LIMIT {limit}"

    return {
        'sql': sql,
        'selected_cols': selected_cols,
        'conditions': conditions,
        'operators': operators,
        'is_distinct': is_distinct,
        'is_count': is_count,
        'count_column': count_column,
        'aggregations': aggregations,
        'order_by': order_by,
        'limit': limit
    }