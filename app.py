import os
import pandas as pd
import dash
from dash import dcc, html
import plotly.express as px

# --- Path to Excel file (Relative path for Render & Local) ---
file_path = "Student.xlsx"

# ✅ Check if Excel file exists
if not os.path.exists(file_path):
    raise FileNotFoundError(f"Excel file not found at: {file_path}")

# --- Load Data ---
df = pd.read_excel(file_path)

# --- Clean column names (remove extra spaces) ---
df.columns = df.columns.str.strip()

# --- Detect Name column ---
name_candidates = ["Name", "Student Name", "NAME", "Names"]
name_col = next((c for c in name_candidates if c in df.columns), None)
if name_col is None:
    raise KeyError(f"No student name column found. Detected: {df.columns.tolist()}")

# --- Detect essential columns ---
required_cols = ["TOTAL", "GRADE"]
for col in required_cols:
    if col not in df.columns:
        raise KeyError(f"Missing required column: {col}")

# --- Identify subject columns ---
exclude_cols = [name_col, "Reg. No", "TOTAL", "AVG.", "MAX", "MIN", "GRADE"]
subject_cols = [c for c in df.columns if c not in exclude_cols]
if not subject_cols:
    raise KeyError("No subject columns detected.")

# --- Melt dataset for subject-wise analysis ---
df_melted = df.melt(
    id_vars=[name_col, "GRADE"],
    value_vars=subject_cols,
    var_name="Subject",
    value_name="Marks"
)

# --- Initialize Dash App ---
app = dash.Dash(__name__)

# ✅ Subject-wise Animated Marks
fig_subjects = px.bar(
    df_melted,
    x=name_col,
    y="Marks",
    color="Subject",
    animation_frame="Subject",
    animation_group=name_col,
    title="📈 Subject-wise Marks (Animated)",
    barmode="group"
)

# ✅ Animated Total Marks by Grade
fig_total = px.bar(
    df,
    x=name_col,
    y="TOTAL",
    color="GRADE",
    animation_frame="GRADE",
    animation_group=name_col,
    title="🏆 Total Marks by Student (Animated by Grade)"
)

# ✅ Grade Distribution Pie
fig_grade = px.pie(
    df,
    names="GRADE",
    title="🎯 Grade Distribution",
    hole=0.3
)

# ✅ Animated Scatterplot: Subject vs Total
df_scatter = df_melted.merge(df[[name_col, "TOTAL"]], on=name_col)
fig_scatter = px.scatter(
    df_scatter,
    x="Marks",
    y="TOTAL",
    color="Subject",
    size="Marks",
    hover_name=name_col,
    animation_frame="Subject",
    animation_group=name_col,
    title="🔗 Subject Contribution to Total Marks"
)

# --- Dashboard Layout ---
app.layout = html.Div([
    html.H1("📊 Students Performance Dashboard - 2025", style={
        "textAlign": "center",
        "color": "#2c3e50",
        "padding": "10px"
    }),
    html.Div([
        dcc.Graph(figure=fig_subjects),
        dcc.Graph(figure=fig_total),
        dcc.Graph(figure=fig_grade),
        dcc.Graph(figure=fig_scatter),
    ], style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "20px"})
])

# --- Run App ---
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))  # ✅ Use Render's dynamic port
    app.run(host="0.0.0.0", port=port, debug=True)
