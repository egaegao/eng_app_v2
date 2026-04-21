import plotly.graph_objects as go
import pandas as pd
from dashboard_config import PLAN_COLOR, ACTUAL_COLOR, GRID_COLOR, BG_COLOR

# ==========================================
# PENGATURAN FONT GLOBAL & HELPER FORMAT
# ==========================================
BASE_FONT = dict(size=13, color="black")       # Untuk angka di sumbu X, Y, dan Legend
BAR_TEXT_FONT = dict(size=14, color="black")   # Untuk teks/nilai yang melayang di atas Bar
TITLE_FONT = dict(size=16, color="black")      # Untuk Judul Grafik

def format_idn_number(x):
    """Format angka ke standar Indonesia: 1.234,5"""
    try:
        # Menggunakan placeholder 'X' agar tidak terjadi konflik saat replace
        return f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return x

def plot_snapshot_chart(sub_df, metric_label):
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        name="Plan", x=sub_df["metric"], y=sub_df["plan"],
        text=sub_df["plan"].apply(format_idn_number), textposition="outside", marker_color=PLAN_COLOR,
        textfont=BAR_TEXT_FONT, cliponaxis=False,
        hovertemplate="<b>Plan</b><br>%{x}<br>%{y:,.1f}<extra></extra>"
    ))
    
    fig.add_trace(go.Bar(
        name="Actual", x=sub_df["metric"], y=sub_df["actual"],
        text=sub_df["actual"].apply(format_idn_number), textposition="outside", marker_color=ACTUAL_COLOR,
        textfont=BAR_TEXT_FONT, cliponaxis=False,
        hovertemplate="<b>Actual</b><br>%{x}<br>%{y:,.1f}<extra></extra>"
    ))
    
    fig.update_layout(
        title=dict(text=f"<b>{metric_label}</b>", x=0.02, xanchor="left", font=TITLE_FONT),
        barmode="group",
        height=340, margin=dict(l=10, r=10, t=60, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=BASE_FONT),
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR,
        font=BASE_FONT
    )
    fig.update_xaxes(title=None, showgrid=False, tickfont=BASE_FONT)
    fig.update_yaxes(title=None, showgrid=True, gridcolor=GRID_COLOR, zeroline=False, tickfont=BASE_FONT)
    
    return fig

def plot_trend_chart(sub_df, metric_label):
    chart_df = sub_df.copy().sort_values("week_date")
    chart_df["week_label"] = chart_df["week_date"].dt.strftime("%d %b")

    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=chart_df["week_label"], y=chart_df["actual"], mode="lines+markers",
        name="Actual", line=dict(color=ACTUAL_COLOR, width=3, shape="spline", smoothing=0.6),
        marker=dict(size=8), 
        hovertemplate="<b>Actual</b><br>%{x}<br>%{y:,.1f}<extra></extra>"
    ))
    
    fig.add_trace(go.Scatter(
        x=chart_df["week_label"], y=chart_df["plan"], mode="lines+markers",
        name="Plan", line=dict(color="#9AA5B1", width=2, dash="dash", shape="spline", smoothing=0.6),
        marker=dict(size=6), 
        hovertemplate="<b>Plan</b><br>%{x}<br>%{y:,.1f}<extra></extra>"
    ))
    
    fig.update_layout(
        title=dict(text=f"<b>Trend: {metric_label}</b>", x=0.02, xanchor="left", font=TITLE_FONT),
        height=420, 
        margin=dict(l=10, r=10, t=60, b=100), 
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=BASE_FONT),
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR, hovermode="x unified",
        font=BASE_FONT
    )
    
    fig.update_xaxes(
        title=None, 
        showgrid=False, 
        tickfont=BASE_FONT,
        tickangle=-90,      
        dtick=1,            
        automargin=True      
    )
    
    fig.update_yaxes(title=None, showgrid=True, gridcolor=GRID_COLOR, zeroline=False, tickfont=BASE_FONT)
    
    return fig

def plot_simple_bar(df, x_col, plan_col, actual_col, title, y_title):
    fig = go.Figure()

    # ==========================================
    # CUSTOM COLOR KHUSUS FUEL (MODERN STYLE)
    # ==========================================
    if "Fuel" in title:
        plan_color = "#D847AD"    # Indigo (modern)
        actual_color = "#63DFB5"  # Emerald (clean & fresh)
    else:
        plan_color = PLAN_COLOR
        actual_color = ACTUAL_COLOR
    
    fig.add_trace(go.Bar(
        name="Plan", x=df[x_col], y=df[plan_col],
        marker_color=plan_color, 
        text=df[plan_col].apply(format_idn_number), textposition="outside",
        textfont=BAR_TEXT_FONT, cliponaxis=False
    ))
    
    fig.add_trace(go.Bar(
        name="Actual", x=df[x_col], y=df[actual_col],
        marker_color=actual_color, 
        text=df[actual_col].apply(format_idn_number), textposition="outside",
        textfont=BAR_TEXT_FONT, cliponaxis=False
    ))
    
    fig.update_layout(
        title=dict(text=f"<b>{title}</b>", x=0.02, xanchor="left", font=TITLE_FONT), 
        barmode="group",
        height=360, margin=dict(l=10, r=10, t=60, b=10),
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=BASE_FONT),
        font=BASE_FONT
    )
    fig.update_xaxes(title=None, showgrid=False, tickfont=BASE_FONT)
    fig.update_yaxes(title=dict(text=f"<b>{y_title}</b>", font=BASE_FONT), showgrid=True, gridcolor=GRID_COLOR, zeroline=False, tickfont=BASE_FONT)
    
    return fig

def plot_inventory_bar(df):
    fig = go.Figure()
    
    for inv_type in df["inventory_type"].dropna().unique():
        sub = df[df["inventory_type"] == inv_type]
        fig.add_trace(go.Bar(
            name=str(inv_type), x=sub["block"], y=sub["volume_ton"],
            text=sub["volume_ton"].apply(lambda x: f"{int(x):,}".replace(",", ".")), 
            textposition="outside",
            textfont=BAR_TEXT_FONT, cliponaxis=False
        ))
        
    fig.update_layout(
        title=dict(text="<b>Inventory by Type</b>", x=0.02, xanchor="left", font=TITLE_FONT), 
        barmode="group",
        height=360, margin=dict(l=10, r=10, t=60, b=10),
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=BASE_FONT),
        font=BASE_FONT
    )
    fig.update_xaxes(title=None, showgrid=False, tickfont=BASE_FONT)
    fig.update_yaxes(title=dict(text="<b>Volume (Ton)</b>", font=BASE_FONT), showgrid=True, gridcolor=GRID_COLOR, zeroline=False, tickfont=BASE_FONT)
    
    return fig