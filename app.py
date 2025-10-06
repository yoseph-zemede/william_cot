import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# ---------- Load Data ----------
file_path = "gold_cot_data.xlsx"
df = pd.read_excel(file_path)

# ---------- Prepare Columns ----------
col_q = df.columns[16]  # 17th column
col_b = df.columns[17]   # Column B
col_c = df.columns[18]   # Column C
# col_h = df.columns[7]   # Column H

# ---------- Clean and Process Data ----------
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df = df.dropna(subset=['Date', col_q, col_b, col_c])
df = df.sort_values('Date')

# Scale Column Q to percentage
df[col_q] = df[col_q] * 100

# ---------- Function to Create Figure ----------
def create_line_chart(x, y, y_title, title, y_format=".1f", y_dtick=None):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=x,
        y=y,
        mode='lines+markers',
        name=y_title,
        line=dict(color='royalblue'),
        marker=dict(size=6)
    ))

    # Initial date range
    start_date = x.iloc[0]
    end_index = min(30, len(x)-1)
    end_date = x.iloc[end_index]

    fig.update_layout(
        xaxis=dict(
            title='Date',
            type='date',
            range=[start_date, end_date],
            rangeslider=dict(visible=True),
            tickangle=45,
            tickformat='%b %d\n%Y',
            ticklabelmode="period"
        ),
        yaxis=dict(
            title=y_title,
            tickformat=y_format,
            dtick=y_dtick
        ),
        title=title,
        hovermode='x unified',
        height=600,
        margin=dict(l=40, r=40, t=60, b=100)
    )
    return fig

# ---------- Streamlit UI ----------
st.set_page_config(layout="wide", page_title="Open Interest Dashboard")

st.title("📊 COT Report Visualization Dashboard")
st.markdown("Select a tab below to view different time series graphs from the Commitments of Trade Report.")

# ---------- Tabs ----------
tab1, tab2, tab3 = st.tabs(["📈 OI %", "🟩 William Long", "🟥 Whilliam Short"])

with tab1:
    st.subheader(f"Open Interest Stochastic Index as Percentage")
    fig1 = create_line_chart(df['Date'], df[col_q], f"Open Interest Stochastic Index as Percentage", f"Open Interest Stochastic Index as Percentage Over Time", y_format=".1f", y_dtick=5)
    st.plotly_chart(fig1, use_container_width=True)

with tab2:
    st.subheader("Commercial Long (Column B)")
    fig2 = create_line_chart(df['Date'], df[col_b], "Commercial Long", "Commercial Long Over Time", y_format=".0f", y_dtick=25000)
    st.plotly_chart(fig2, use_container_width=True)

with tab3:
    st.subheader("Commercial Short (Column C)")
    fig3 = create_line_chart(df['Date'], df[col_c], "Commercial Short", "Commercial Short Over Time", y_format=".0f", y_dtick=25000)
    st.plotly_chart(fig3, use_container_width=True)



# Optional: Footer or credits
st.markdown("---")
st.markdown("📁 Data Source: Commitments of Trade Report Excel File")