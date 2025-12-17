import json
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import itertools

# Pastel color palette
PASTEL_COLORS = px.colors.qualitative.Pastel

def abbreviate_label(label):
    if not isinstance(label, str):
        return str(label)
    label = label.strip().replace("_", " ")
    words = label.split()
    if len(words) > 1:
        return "".join(w[0].upper() for w in words)
    return label[:3].upper() if len(label) > 5 else label.capitalize()

def shorten_series(series):
    if pd.api.types.is_numeric_dtype(series):
        return series
    return pd.Series(series).astype(str).apply(abbreviate_label)

def plot_chart(df, spec, color_iter):
    cols = [c for c in spec[:-1] if str(c).lower() != "none"]
    chart_type = spec[-1].lower().strip()

    if chart_type == "pie":
        col = cols[0]
        vc = df[col].value_counts()
        colors = [next(color_iter) for _ in range(len(vc))]
        trace = go.Pie(
            labels=shorten_series(vc.index),
            values=vc.values,
            hole=0.4,
            marker=dict(colors=colors),
            showlegend=True,
            title=abbreviate_label(col)
        )
        return trace, "domain", abbreviate_label(col)

    elif chart_type == "histogram":
        col = cols[0]
        color = next(color_iter)
        trace = go.Histogram(
            x=shorten_series(df[col]),
            nbinsx=20,
            marker=dict(color=color),
            name=abbreviate_label(col),
            showlegend=True
        )
        return trace, "xy", abbreviate_label(col)

    elif chart_type == "bar":
        if len(cols) == 1:
            col = cols[0]
            vc = df[col].value_counts()
            traces = []
            for val in vc.index:
                traces.append(go.Bar(
                    x=[abbreviate_label(val)],
                    y=[vc[val]],
                    marker=dict(color=next(color_iter)),
                    name=abbreviate_label(val),
                    showlegend=True
                ))
            return traces, "xy", abbreviate_label(col)
        else:
            col1, col2 = cols[0], cols[1]
            if pd.api.types.is_numeric_dtype(df[col2]):
                grouped = df.groupby(col1)[col2].mean().reset_index()
                trace = go.Bar(
                    x=shorten_series(grouped[col1]),
                    y=grouped[col2],
                    marker=dict(color=next(color_iter)),
                    name=f"{abbreviate_label(col1)} vs {abbreviate_label(col2)}",
                    showlegend=True
                )
            else:
                grouped = df.groupby([col1, col2]).size().reset_index(name='count')
                trace = go.Bar(
                    x=shorten_series(grouped[col1]),
                    y=grouped['count'],
                    marker=dict(color=next(color_iter)),
                    name=f"{abbreviate_label(col1)} vs {abbreviate_label(col2)}",
                    showlegend=True
                )
            return trace, "xy", abbreviate_label(col1)

    elif chart_type == "line" and len(cols) >= 2:
        trace = go.Scatter(
            x=shorten_series(df[cols[0]]),
            y=df[cols[1]],
            mode="lines+markers",
            line=dict(color=next(color_iter)),
            name=f"{abbreviate_label(cols[0])} vs {abbreviate_label(cols[1])}",
            showlegend=True
        )
        return trace, "xy", abbreviate_label(cols[0])

    elif chart_type == "scatter" and len(cols) >= 2:
        trace = go.Scatter(
            x=shorten_series(df[cols[0]]),
            y=shorten_series(df[cols[1]]),
            mode="markers",
            marker=dict(color=next(color_iter)),
            name=f"{abbreviate_label(cols[0])} vs {abbreviate_label(cols[1])}",
            showlegend=True
        )
        return trace, "xy", abbreviate_label(cols[0])

    elif chart_type == "area" and len(cols) >= 2:
        trace = go.Scatter(
            x=shorten_series(df[cols[0]]),
            y=df[cols[1]],
            fill="tozeroy",
            line=dict(color=next(color_iter)),
            name=f"{abbreviate_label(cols[0])} vs {abbreviate_label(cols[1])}",
            showlegend=True
        )
        return trace, "xy", abbreviate_label(cols[0])

    return go.Indicator(mode="number", value=0), "indicator", "NA"

def determine_grid(n):
    return (1,1) if n==1 else (1,2) if n==2 else (1,3) if n==3 else (2,2) if n==4 else (2,3)

def create_dashboard(df, json_file="charts.json", output="dashboard.png"):
    with open(json_file) as f:
        charts = json.load(f).get("dashboards", [])

    n = len(charts)
    nrows, ncols = determine_grid(n)

    # Safe grid specs
    specs = []
    for r in range(nrows):
        row_specs = []
        for c in range(ncols):
            idx = r * ncols + c
            if idx < n:
                chart_type = charts[idx][-1].lower()
                row_specs.append({"type":"domain"} if chart_type=="pie" else {"type":"xy"})
            else:
                row_specs.append(None)
        specs.append(row_specs)

    fig = make_subplots(rows=nrows, cols=ncols, specs=specs,
                        horizontal_spacing=0.08, vertical_spacing=0.18)

    color_iter = itertools.cycle(PASTEL_COLORS)

    for i, spec in enumerate(charts):
        r, c = i//ncols + 1, i%ncols + 1
        traces, subtype, title = plot_chart(df, spec, color_iter)
        if isinstance(traces, list):
            for trace in traces:
                fig.add_trace(trace, row=r, col=c)
        else:
            fig.add_trace(traces, row=r, col=c)

        # Remove axis titles completely
        fig.update_xaxes(title_text=None, row=r, col=c)
        fig.update_yaxes(title_text=None, row=r, col=c)

        # Shorten X ticks if categorical
        if spec[0].lower() != "none" and not pd.api.types.is_numeric_dtype(df[spec[0]]):
            x_vals = shorten_series(df[spec[0]].unique())
            fig.update_xaxes(ticktext=list(x_vals), tickvals=list(range(len(x_vals))), row=r, col=c)

        # Shorten Y ticks if second column exists and categorical
        if len(spec) > 1 and spec[1].lower() != "none" and not pd.api.types.is_numeric_dtype(df[spec[1]]):
            y_vals = shorten_series(df[spec[1]].unique())
            fig.update_yaxes(ticktext=list(y_vals), tickvals=list(range(len(y_vals))), row=r, col=c)

        # Subplot titles for non-pie charts
        # if spec[-1].lower() == "histogram":
        #     fig.add_annotation(
        #         text=f"<b>{title}</b>",
        #         xref="x domain",
        #         yref="y domain",
        #         x=0.5,
        #         y=1.05,
        #         showarrow=False,
        #         font=dict(size=14),
        #         align="center",
        #         row=r, col=c
        #     )

    fig.update_layout(
        height=900 if nrows==2 else 500,
        width=1500,
        title_text="",  # empty dashboard title
        margin=dict(l=60,r=60,t=60,b=80),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5)
    )

    try:
        fig.write_image(output, scale=2)
        print(f"✅ Dashboard saved as {output}")
        return output
    except ValueError:
        print("⚠️ Install 'kaleido' for saving images: pip install kaleido")

if __name__ == "__main__":
    df = pd.read_csv("data.csv")
    create_dashboard(df)