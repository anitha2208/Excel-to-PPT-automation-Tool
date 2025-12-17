# core/query_handler.py
import pandas as pd
import numpy as np
import json
import os
from datetime import datetime
from nlp_to_sql import generate_sql

def load_dataset(file_path: str) -> pd.DataFrame:
    if file_path.endswith(".csv"):
        return pd.read_csv(file_path)
    elif file_path.endswith(".xlsx") or file_path.endswith(".xls"):
        return pd.read_excel(file_path, sheet_name=0)
    else:
        raise ValueError("Unsupported file type. Use CSV or Excel.")

# ---------------- Conditions + Aggregations (same as before) ----------------
def apply_single_condition(df, col, operator, value):
    try:
        if operator == '=':
            if pd.api.types.is_numeric_dtype(df[col]):
                value = float(value) if str(value).replace('.', '', 1).isdigit() else value
                return df[col] == value
            return df[col].astype(str).str.lower() == str(value).lower()
        elif operator == '!=':
            return df[col].astype(str).str.lower() != str(value).lower()
        elif operator in ['>', '<', '>=', '<=']:
            numeric_col = pd.to_numeric(df[col], errors='coerce')
            value = float(value)
            return eval(f"numeric_col {operator} value")
        elif operator == 'BETWEEN':
            numeric_col = pd.to_numeric(df[col], errors='coerce')
            return (numeric_col >= value[0]) & (numeric_col <= value[1])
        elif operator == 'STARTS_WITH':
            return df[col].astype(str).str.lower().str.startswith(str(value).lower())
        elif operator == 'ENDS_WITH':
            return df[col].astype(str).str.lower().str.endswith(str(value).lower())
        elif operator == 'CONTAINS':
            return df[col].astype(str).str.lower().str.contains(str(value).lower(), na=False)
        return pd.Series([True] * len(df))
    except:
        return pd.Series([True] * len(df))

def apply_conditions(df, conditions, operators):
    if not conditions:
        return df
    mask = apply_single_condition(df, *conditions[0])
    for i, cond in enumerate(conditions[1:], 1):
        next_mask = apply_single_condition(df, *cond)
        op = operators[i-1] if i-1 < len(operators) else "AND"
        mask = mask & next_mask if op == "AND" else mask | next_mask
    return df[mask]

def apply_aggregations(df, aggregations):
    results = {}
    for col, func in aggregations:
        numeric_col = pd.to_numeric(df[col], errors='coerce')
        if func == 'sum':
            results[f"sum_{col}"] = numeric_col.sum()
        elif func == 'mean':
            results[f"avg_{col}"] = numeric_col.mean()
        elif func == 'min':
            results[f"min_{col}"] = numeric_col.min()
        elif func == 'max':
            results[f"max_{col}"] = numeric_col.max()
        elif func == 'count':
            results[f"count_{col}"] = numeric_col.count()
    return results

# ---------------- Query Executor ----------------
def execute_query(file_path: str, user_query: str):
    df = load_dataset(file_path)
    query_info = generate_sql(user_query, df)

    filtered_df = apply_conditions(df, query_info['conditions'], query_info['operators'])

    if query_info['aggregations']:
        result_data = [apply_aggregations(filtered_df, query_info['aggregations'])]
    elif query_info['is_count']:
        result_data = [{"count": len(filtered_df)}]
    else:
        result_df = filtered_df[query_info['selected_cols']] if query_info['selected_cols'] else filtered_df
        if query_info['is_distinct']:
            result_df = result_df.drop_duplicates()
        if query_info['order_by']:
            col, direction = query_info['order_by']
            result_df = result_df.sort_values(by=col, ascending=(direction=="ASC"))
        if query_info['limit']:
            result_df = result_df.head(query_info['limit'])
        result_data = result_df.to_dict('records')

    outputs_dir = "output"
    os.makedirs(outputs_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(outputs_dir,f"query_output_{ts}.json")

    def convert(obj):
        if isinstance(obj, (np.integer, np.int64)): return int(obj)
        if isinstance(obj, (np.floating, np.float64)): return float(obj)
        if pd.isna(obj): return None
        return obj

    json_data = [{k: convert(v) for k, v in record.items()} for record in result_data]
    with open(out_path, "w") as f:
        json.dump(json_data, f, indent=2)
    return out_path