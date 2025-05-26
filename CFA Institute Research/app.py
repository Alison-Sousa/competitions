import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Set page config
st.set_page_config(
    page_title="Vivara Equity Research",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main {
        background-color: #f8f9fa;
    }
    .st-emotion-cache-1v0mbdj {
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .header-style {
        font-size: 24px;
        font-weight: bold;
        color: #2c3e50;
        margin-bottom: 10px;
    }
    .subheader-style {
        font-size: 18px;
        font-weight: bold;
        color: #3498db;
        margin-bottom: 10px;
    }
    .metric-box {
        background-color: white;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
    }
    .positive {
        color: #27ae60;
        font-weight: bold;
    }
    .negative {
        color: #e74c3c;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# Header
col1, col2 = st.columns([3, 1])
with col1:
    st.title("Vivara Participações S.A. (VIVA3) - Equity Research")
    st.markdown("**Brazil's Leading Jewelry Retailer - Initiating Coverage with BUY Recommendation**")
with col2:
    st.image("https://wp-cdn.etiquetaunica.com.br/blog/wp-content/uploads/2019/11/24101844/capa-post-sobre-a-vivara-241119.jpg", width=150)

# Current Price and Target
current_price = 25.60
target_price = 32.83
upside = ((target_price - current_price) / current_price) * 100

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Current Price (BRL)", f"{current_price:.2f}")
with col2:
    st.metric("12M Target Price (BRL)", f"{target_price:.2f}")
with col3:
    st.metric("Upside Potential", f"{upside:.1f}%", delta=f"{upside:.1f}%")
with col4:
    st.metric("Recommendation", "BUY", delta="Strong Buy")

# Monte Carlo Simulation
st.markdown('<div class="header-style">Monte Carlo Simulation</div>', unsafe_allow_html=True)

# Data for Monte Carlo
np.random.seed(42)
simulations = 10000
mean_return = 0.28  # 28% expected return
volatility = 0.15  # 15% volatility
simulated_prices = current_price * (1 + np.random.normal(mean_return, volatility, simulations) / 252 * 365)

# Calculate probabilities
buy_threshold = target_price * 0.95
hold_threshold = target_price * 0.85
sell_threshold = target_price * 0.75

buy_prob = np.mean(simulated_prices >= buy_threshold) * 100
hold_prob = np.mean((simulated_prices >= hold_threshold) & (simulated_prices < buy_threshold)) * 100
sell_prob = np.mean(simulated_prices < hold_threshold) * 100

# Create figure
fig = go.Figure()

# Add histogram
fig.add_trace(go.Histogram(
    x=simulated_prices,
    nbinsx=50,
    marker_color='#3498db',
    opacity=0.7,
    name='Price Distribution'
))

# Add vertical lines
fig.add_vline(x=current_price, line_dash="dash", line_color="green", annotation_text=f"Current: BRL {current_price}", 
              annotation_position="top right")
fig.add_vline(x=target_price, line_dash="dash", line_color="red", annotation_text=f"Target: BRL {target_price}", 
              annotation_position="top right")

# Update layout
fig.update_layout(
    title="Monte Carlo Simulation of 12M Price Target",
    xaxis_title="Price (BRL)",
    yaxis_title="Frequency",
    bargap=0.1,
    height=500,
    showlegend=False
)

# Display chart and probabilities
col1, col2 = st.columns([3, 1])
with col1:
    st.plotly_chart(fig, use_container_width=True)
with col2:
    st.markdown('<div class="metric-box">', unsafe_allow_html=True)
    st.metric("Probability of BUY", f"{69}%", delta="69%", delta_color="normal")
    st.metric("Probability of HOLD", f"{22}%", delta="22%", delta_color="normal")
    st.metric("Probability of SELL", f"{9}%", delta="9%", delta_color="normal")
    st.markdown('</div>', unsafe_allow_html=True)

# Investment Highlights
st.markdown('<div class="header-style">Investment Highlights</div>', unsafe_allow_html=True)

highlight_col1, highlight_col2, highlight_col3 = st.columns(3)

with highlight_col1:
    st.markdown('<div class="metric-box">', unsafe_allow_html=True)
    st.markdown('<div class="subheader-style">Resilient Industry</div>', unsafe_allow_html=True)
    st.markdown("""
    - Brazilian jewelry sector CAGR: 4.6% (2007-2019)
    - Market size: BRL 12.6bn (2021)
    - Top 4 players: 22.3% market share
    - Vivara market share: 16% (up from 11.4% in 2019)
    """)
    st.markdown('</div>', unsafe_allow_html=True)

with highlight_col2:
    st.markdown('<div class="metric-box">', unsafe_allow_html=True)
    st.markdown('<div class="subheader-style">Business Model Advantages</div>', unsafe_allow_html=True)
    st.markdown("""
    - Vertical integration (Manaus factory)
    - Strong brand awareness (79% top-of-mind)
    - Scale advantages (38% of Brazilian malls)
    - Tax benefits (25% of net income)
    """)
    st.markdown('</div>', unsafe_allow_html=True)

with highlight_col3:
    st.markdown('<div class="metric-box">', unsafe_allow_html=True)
    st.markdown('<div class="subheader-style">Growth Potential</div>', unsafe_allow_html=True)
    st.markdown("""
    - 91 potential Vivara store openings
    - 259 potential Life store openings
    - eCommerce penetration: 8% (vs 3% pre-pandemic)
    - Life stores ROIC: 46% (vs 22-35% for Vivara)
    """)
    st.markdown('</div>', unsafe_allow_html=True)

# Financial Analysis
st.markdown('<div class="header-style">Financial Analysis</div>', unsafe_allow_html=True)

# Financial Metrics Data
years = ['2016', '2017', '2018', '2019', '2020', '2021', '2022E', '2023E', '2024E', '2025E']
net_revenue = [1466, 1590, 2046, 2619, 3326, 3956, 4495, 4955, 5365, 5656]  # in BRL millions
ebitda = [286, 256, 377, 543, 756, 959, 1137, 1295, 1437, 1540]  # in BRL millions
net_income = [298, 245, 336, 464, 642, 837, 1013, 1166, 1302, 1404]  # in BRL millions
gross_margin = [67.6, 67.2, 68.0, 69.2, 70.0, 70.5, 70.9, 71.2, 71.5, 71.6]  # %
ebitda_margin = [19.5, 16.1, 18.4, 20.7, 22.7, 24.2, 25.3, 26.1, 26.6, 26.9]  # %
net_margin = [20.4, 15.4, 16.4, 17.7, 19.3, 21.1, 22.5, 23.5, 24.1, 24.5]  # %
roic = [35.2, 17.1, 16.1, 17.2, 19.5, 22.5, 24.9, 26.5, 27.7, 28.6]  # %
stores = [267, 328, 416, 500, 558, 578, 598, 618, 625, 625]

# Create financial metrics dataframe
financial_df = pd.DataFrame({
    'Year': years,
    'Net Revenue (BRL mn)': net_revenue,
    'EBITDA (BRL mn)': ebitda,
    'Net Income (BRL mn)': net_income,
    'Gross Margin (%)': gross_margin,
    'EBITDA Margin (%)': ebitda_margin,
    'Net Margin (%)': net_margin,
    'ROIC (%)': roic,
    'Number of Stores': stores
})

# Display financial metrics
col1, col2 = st.columns([1, 3])
with col1:
    st.dataframe(financial_df.set_index('Year'), height=600)

with col2:
    # Create subplots
    fig = make_subplots(rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.1)

    # Revenue, EBITDA, Net Income
    fig.add_trace(go.Bar(
        x=years,
        y=net_revenue,
        name='Net Revenue (BRL mn)',
        marker_color='#3498db'
    ), row=1, col=1)

    fig.add_trace(go.Scatter(
        x=years,
        y=ebitda,
        name='EBITDA (BRL mn)',
        line=dict(color='#2ecc71', width=3)
    ), row=2, col=1)

    fig.add_trace(go.Scatter(
        x=years,
        y=net_income,
        name='Net Income (BRL mn)',
        line=dict(color='#9b59b6', width=3)
    ), row=3, col=1)

    # Update layout
    fig.update_layout(
        height=600,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    fig.update_yaxes(title_text="Net Revenue (BRL mn)", row=1, col=1)
    fig.update_yaxes(title_text="EBITDA (BRL mn)", row=2, col=1)
    fig.update_yaxes(title_text="Net Income (BRL mn)", row=3, col=1)

    st.plotly_chart(fig, use_container_width=True)

# Margins Chart
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=years,
    y=gross_margin,
    name='Gross Margin',
    line=dict(color='#3498db', width=3)
))

fig.add_trace(go.Scatter(
    x=years,
    y=ebitda_margin,
    name='EBITDA Margin',
    line=dict(color='#2ecc71', width=3)
))

fig.add_trace(go.Scatter(
    x=years,
    y=net_margin,
    name='Net Margin',
    line=dict(color='#9b59b6', width=3)
))

fig.update_layout(
    title="Margin Evolution (%)",
    xaxis_title="Year",
    yaxis_title="Margin (%)",
    height=400,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)

st.plotly_chart(fig, use_container_width=True)

# Valuation Analysis
st.markdown('<div class="header-style">Valuation Analysis</div>', unsafe_allow_html=True)

# DCF Assumptions
st.markdown('<div class="subheader-style">DCF Assumptions</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1:
    st.markdown('<div class="metric-box">', unsafe_allow_html=True)
    st.markdown("**Cost of Equity**")
    st.markdown("- Risk-Free Rate: 4.0%")
    st.markdown("- Beta: 1.0")
    st.markdown("- Equity Risk Premium: 5.1%")
    st.markdown("- Brazil Risk Premium: 3.6%")
    st.markdown("- **Ke: 13.8%**")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="metric-box">', unsafe_allow_html=True)
    st.markdown("**Cost of Debt**")
    st.markdown("- Pre-tax Kd: 12.6%")
    st.markdown("- After-tax Kd: 10.7%")
    st.markdown("**WACC**")
    st.markdown("- **13.1%**")
    st.markdown('</div>', unsafe_allow_html=True)

with col3:
    st.markdown('<div class="metric-box">', unsafe_allow_html=True)
    st.markdown("**Terminal Growth**")
    st.markdown("- Brazil LT Inflation: 3.5%")
    st.markdown("- Real Growth: 1.0%")
    st.markdown("- **Terminal Growth: 4.5%**")
    st.markdown('</div>', unsafe_allow_html=True)

# DCF Results
st.markdown('<div class="subheader-style">DCF Valuation</div>', unsafe_allow_html=True)

dcf_col1, dcf_col2 = st.columns([1, 2])

with dcf_col1:
    st.markdown('<div class="metric-box">', unsafe_allow_html=True)
    st.metric("Enterprise Value (BRL mn)", "7,532")
    st.metric("Equity Value (BRL mn)", "7,754")
    st.metric("Target Price (BRL)", "32.83")
    st.metric("Upside Potential", f"{upside:.1f}%", delta=f"{upside:.1f}%")
    st.markdown('</div>', unsafe_allow_html=True)

with dcf_col2:
    # FCFF Projection
    years_fcff = ['2H22E', '2023E', '2024E', '2025E', '2026E', '2027E', '2028E', '2029E', '2030E', '2031E', '2032E']
    fcff = [-80, -188, -102, 93, 476, 719, 922, 1133, 1324, 1438, 1518]  # in BRL millions
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=years_fcff,
        y=fcff,
        name='FCFF (BRL mn)',
        marker_color=np.where(np.array(fcff) > 0, '#2ecc71', '#e74c3c')
    ))
    
    fig.update_layout(
        title="Free Cash Flow to Firm Projection (BRL mn)",
        xaxis_title="Year",
        yaxis_title="FCFF (BRL mn)",
        height=400
    )
    
    st.plotly_chart(fig, use_container_width=True)

# Relative Valuation
st.markdown('<div class="subheader-style">Relative Valuation</div>', unsafe_allow_html=True)

# Trading Comps Data
companies = ['Vivara', 'Grupo Soma', 'Arezzo', 'LVMH', 'Pandora', 'Signet Jewelers']
ev_ebitda_2023 = [12.7, 12.9, 13.8, 21.1, 7.3, 6.4]
ebitda_growth = [18.6, 19.2, 21.1, 21.1, 7.3, 6.4]
roic = [25.8, 8.4, 20.4, 18.9, 39.7, 25.0]

comp_df = pd.DataFrame({
    'Company': companies,
    '2023E EV/EBITDA': ev_ebitda_2023,
    'EBITDA Growth (21-24E)': ebitda_growth,
    'ROIC (%)': roic
})

# Display comps
col1, col2 = st.columns([1, 2])
with col1:
    st.dataframe(comp_df.set_index('Company'))

with col2:
    fig = px.scatter(
        comp_df, 
        x='EBITDA Growth (21-24E)', 
        y='2023E EV/EBITDA',
        size='ROIC (%)',
        color='Company',
        hover_name='Company',
        title='EV/EBITDA vs. EBITDA Growth (Size = ROIC)',
        height=500
    )
    
    # Add regression line
    m, b = np.polyfit(ebitda_growth, ev_ebitda_2023, 1)
    fig.add_trace(go.Scatter(
        x=np.array(ebitda_growth),
        y=m * np.array(ebitda_growth) + b,
        mode='lines',
        line=dict(color='gray', dash='dash'),
        name='Regression Line'
    ))
    
    st.plotly_chart(fig, use_container_width=True)

# Risk Analysis
st.markdown('<div class="header-style">Risk Analysis</div>', unsafe_allow_html=True)

# Risk Matrix Data
risks = [
    {"Risk": "Acquisition & Integration", "Type": "Business", "Impact": "High", "Probability": "Medium"},
    {"Risk": "Cannibalization", "Type": "Business", "Impact": "Medium", "Probability": "High"},
    {"Risk": "Tax Benefit Loss", "Type": "Business", "Impact": "High", "Probability": "Low"},
    {"Risk": "Management Quality", "Type": "Business", "Impact": "Medium", "Probability": "Medium"},
    {"Risk": "Peers Capitulation", "Type": "Market", "Impact": "Medium", "Probability": "Low"},
    {"Risk": "International Competition", "Type": "Market", "Impact": "Medium", "Probability": "Medium"},
    {"Risk": "Mall Vacancy Rates", "Type": "Market", "Impact": "Low", "Probability": "Medium"},
    {"Risk": "Tax Reform", "Type": "Market", "Impact": "High", "Probability": "Medium"},
    {"Risk": "Mall Relevance Decline", "Type": "Market", "Impact": "High", "Probability": "Low"},
    {"Risk": "Macro Crisis", "Type": "Economic", "Impact": "High", "Probability": "Medium"},
    {"Risk": "Consumer Preferences", "Type": "Economic", "Impact": "Medium", "Probability": "Medium"},
]

risk_df = pd.DataFrame(risks)

# Impact and Probability mapping
impact_map = {"Low": 1, "Medium": 2, "High": 3}
prob_map = {"Low": 1, "Medium": 2, "High": 3}

risk_df['Impact Score'] = risk_df['Impact'].map(impact_map)
risk_df['Probability Score'] = risk_df['Probability'].map(prob_map)

# Create risk matrix
fig = px.scatter(
    risk_df,
    x='Probability Score',
    y='Impact Score',
    color='Type',
    symbol='Type',
    hover_name='Risk',
    title='Risk Matrix',
    height=500,
    color_discrete_map={
        "Business": "#3498db",
        "Market": "#2ecc71",
        "Economic": "#9b59b6"
    }
)

# Add risk quadrants
fig.add_shape(type="rect", x0=0.5, y0=2.5, x1=3.5, y1=3.5, line=dict(color="Red", width=2), fillcolor="rgba(255,0,0,0.1)")
fig.add_shape(type="rect", x0=2.5, y0=1.5, x1=3.5, y1=2.5, line=dict(color="Orange", width=2), fillcolor="rgba(255,165,0,0.1)")
fig.add_shape(type="rect", x0=0.5, y0=1.5, x1=2.5, y1=2.5, line=dict(color="Yellow", width=2), fillcolor="rgba(255,255,0,0.1)")
fig.add_shape(type="rect", x0=0.5, y0=0.5, x1=1.5, y1=1.5, line=dict(color="Green", width=2), fillcolor="rgba(0,255,0,0.1)")

# Add quadrant labels
fig.add_annotation(x=1, y=3, text="Critical Risks", showarrow=False, font=dict(color="red"))
fig.add_annotation(x=3, y=2, text="Major Risks", showarrow=False, font=dict(color="orange"))
fig.add_annotation(x=1.5, y=2, text="Moderate Risks", showarrow=False, font=dict(color="goldenrod"))
fig.add_annotation(x=1, y=1, text="Minor Risks", showarrow=False, font=dict(color="green"))

# Update axes
fig.update_xaxes(tickvals=[1, 2, 3], ticktext=["Low", "Medium", "High"])
fig.update_yaxes(tickvals=[1, 2, 3], ticktext=["Low", "Medium", "High"])

# Display risk matrix
st.plotly_chart(fig, use_container_width=True)

# Scenario Analysis
st.markdown('<div class="subheader-style">Scenario Analysis</div>', unsafe_allow_html=True)

scenarios = ["Bull", "Base", "Bear"]
sss = [10.0, 9.3, 6.0]  # %
tax_benefit = ["Yes", "Yes", "No"]
cannibalization = [7.0, 12.5, 23.0]  # %
life_openings = [157, 142, 97]
target_prices = [37.7, 32.8, 21.65]
upsides = [47.2, 28.8, -15.7]  # %

scenario_df = pd.DataFrame({
    "Scenario": scenarios,
    "SSS Life (YoY %)": sss,
    "Tax Benefit Renewal": tax_benefit,
    "Cannibalization Rate (%)": cannibalization,
    "Life Openings Cluster C": life_openings,
    "Target Price (BRL)": target_prices,
    "Upside/Downside (%)": upsides
})

# Display scenario table
st.dataframe(scenario_df.set_index('Scenario'))

# Cannibalization Sensitivity
st.markdown('<div class="subheader-style">Cannibalization Sensitivity Analysis</div>', unsafe_allow_html=True)

# Create heatmap data
cannibalization_rates = [0, 20, 40, 60, 80, 100]  # %
wacc_rates = [0, 20, 40, 60, 80, 100]  # %
data = np.array([
    [31.7, 27.8, 23.9, 20.1, 16.2, 12.3],
    [28.4, 24.5, 20.6, 16.7, 12.8, 8.9],
    [26.0, 21.1, 17.2, 13.3, 9.4, 5.5],
    [21.6, 17.8, 13.9, 10.0, 6.1, 2.2],
    [18.3, 14.4, 10.5, 6.6, 2.7, -1.2],
    [14.9, 11.0, 7.1, 3.2, -0.6, -4.5]
])

# Create heatmap
fig = go.Figure(data=go.Heatmap(
    z=data,
    x=cannibalization_rates,
    y=wacc_rates,
    colorscale='RdYlGn',
    zmid=0,
    colorbar=dict(title="Fair Value (BRL)")
))

fig.update_layout(
    title="Fair Value Sensitivity to Cannibalization and WACC (BRL)",
    xaxis_title="Cannibalization Rate (%)",
    yaxis_title="WACC (%)",
    height=500
)

st.plotly_chart(fig, use_container_width=True)

# Footer
st.markdown("---")
st.markdown("**Equity Research | CFA Institute Research Challenge**")
st.markdown("**Disclaimer:** This report is for informational purposes only and does not constitute investment advice.")