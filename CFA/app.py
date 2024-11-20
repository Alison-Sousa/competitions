import yfinance as yf
import numpy as np
from random import random
import matplotlib.pyplot as plt
from scipy.stats import norm
import streamlit as st
from datetime import datetime, timedelta


dict_name_tick = {'NASDAQ': '^IXIC', 'S&P 500': '^GSPC', 'Vivara Participações S.A.': 'VIVA3.SA', 'Tesla, Inc.': 'TSLA', 'Lyft, Inc.': 'LYFT', 'Banco Bradesco S.A.': 'BBD', 'Ford Motor Company': 'F', 'Amazon.com, Inc.': 'AMZN', 'NVIDIA Corporation': 'NVDA', 'Alphabet Inc.': 'GOOG', 'Apple Inc.': 'AAPL', 'Advanced Micro Devices, Inc.': 'AMD', 'Itaú Unibanco Holding S.A.': 'ITUB', 'Palantir Technologies Inc.': 'PLTR', 'Uber Technologies, Inc.': 'UBER', 'NIO Inc.': 'NIO', 'Carnival Corporation & plc': 'CCL', 'AMC Entertainment Holdings, Inc.': 'APE', 'SoFi Technologies, Inc.': 'SOFI', 'PayPal Holdings, Inc.': 'PYPL', 'Lumen Technologies, Inc.': 'LUMN', 'Petróleo Brasileiro S.A. - Petrobras': 'PBR-A', 'Meta Platforms, Inc.': 'META', 'Bank of America Corporation': 'BAC', 'Intel Corporation': 'INTC', 'Credit Suisse Group AG': 'CS', 'Microsoft Corporation': 'MSFT', 'Snap Inc.': 'SNAP', 'AT&T Inc.': 'T', 'Exxon Mobil Corporation': 'XOM', 'Transocean Ltd.': 'RIG', 'C3.ai, Inc.': 'AI', 'Vale S.A.': 'VALE', 'Lucid Group, Inc.': 'LCID', 'Affirm Holdings, Inc.': 'AFRM', 'Southwestern Energy Company': 'SWN', 'Ambev S.A.': 'ABEV', 'Ginkgo Bioworks Holdings, Inc.': 'DNA', 'Nu Holdings Ltd.': 'NU', 'BP p.l.c.': 'BP', 'Sirius XM Holdings Inc.': 'SIRI', 'Wells Fargo & Company': 'WFC', 'Alibaba Group Holding Limited': 'BABA', 'Cloudflare, Inc.': 'NET', 'Pfizer Inc.': 'PFE', 'iQIYI, Inc.': 'IQ', 'Cisco Systems, Inc.': 'CSCO', 'Coinbase Global, Inc.': 'COIN', 'Rivian Automotive, Inc.': 'RIVN', 'Citigroup Inc.': 'C', 'Comcast Corporation': 'CMCSA', 'AGNC Investment Corp.': 'AGNC', 'Lufax Holding Ltd': 'LU', 'The Walt Disney Company': 'DIS', 'Verizon Communications Inc.': 'VZ', 'Shopify Inc.': 'SHOP', 'DraftKings Inc.': 'DKNG', 'Kinross Gold Corporation': 'KGC', 'Peloton Interactive, Inc.': 'PTON', 'Grab Holdings Limited': 'GRAB', 'American Airlines Group Inc.': 'AAL', 'Hewlett Packard Enterprise Company': 'HPE', 'V.F. Corporation': 'VFC', 'Unity Software Inc.': 'U', 'Occidental Petroleum Corporation': 'OXY', 'Nokia Oyj': 'NOK', 'Pinterest, Inc.': 'PINS', 'Kinder Morgan, Inc.': 'KMI', 'Barrick Gold Corporation': 'GOLD', 'PG&E Corporation': 'PCG', 'XPeng Inc.': 'XPEV', 'Baxter International Inc.': 'BAX', 'The Coca-Cola Company': 'KO', 'Warner Bros. Discovery, Inc.': 'WBD', 'Infosys Limited': 'INFY', 'DiDi Global Inc.': 'DIDIY', 'Chevron Corporation': 'CVX', 'Medical Properties Trust, Inc.': 'MPW', 'Taiwan Semiconductor Manufacturing Company Limited': 'TSM', 'Plug Power Inc.': 'PLUG', 'Micron Technology, Inc.': 'MU', 'General Motors Company': 'GM', 'ChargePoint Holdings, Inc.': 'CHPT', 'Block, Inc.': 'SQ', 'Salesforce, Inc.': 'CRM', 'Farfetch Limited': 'FTCH', 'Newell Brands Inc.': 'NWL', 'JD.com, Inc.': 'JD', 'CSX Corporation': 'CSX', 'Marathon Oil Corporation': 'MRO', 'Gerdau S.A.': 'GGB', 'Roblox Corporation': 'RBLX', 'Fisker Inc.': 'FSR', 'NextEra Energy, Inc.': 'NEE', 'ConocoPhillips': 'COP', 'Energy Transfer LP': 'ET', 'Teva Pharmaceutical Industries Limited': 'TEVA', 'Oak Street Health, Inc.': 'OSH', 'Norwegian Cruise Line Holdings Ltd.': 'NCLH', 'Schlumberger Limited': 'SLB', 'Expedia Group, Inc.': 'EXPE', 'Kosmos Energy Ltd.': 'KOS', 'Coterra Energy Inc.': 'CTRA', 'Annaly Capital Management, Inc.': 'NLY', 'KE Holdings Inc.': 'BEKE', 'AbbVie Inc.': 'ABBV', 'Flex Ltd.': 'FLEX', 'Luminar Technologies, Inc.': 'LAZR', 'Exelon Corporation': 'EXC', 'EQT Corporation': 'EQT', 'Lloyds Banking Group plc': 'LYG', 'Devon Energy Corporation': 'DVN', 'Li Auto Inc.': 'LI', 'Bristol-Myers Squibb Company': 'BMY', 'Robinhood Markets, Inc.': 'HOOD', 'News Corporation': 'NWSA', 'Halliburton Company': 'HAL', 'The Procter & Gamble Company': 'PG', 'Doximity, Inc.': 'DOCS', 'Delta Air Lines, Inc.': 'DAL', 'ASE Technology Holding Co., Ltd.': 'ASX', 'Freeport-McMoRan Inc.': 'FCX', 'APA Corporation': 'APA', 'Bilibili Inc.': 'BILI', 'KeyCorp': 'KEY', 'Paramount Global': 'PARA', 'Cleveland-Cliffs Inc.': 'CLF', 'CVS Health Corporation': 'CVS', 'Huntington Bancshares Incorporated': 'HBAN', 'Marvell Technology, Inc.': 'MRVL', 'New York Community Bancorp, Inc.': 'NYCB', 'Avantor, Inc.': 'AVTR', 'Under Armour, Inc.': 'UAA', 'The Western Union Company': 'WU', 'Shell plc': 'SHEL', "Macy's, Inc.": 'M', 'Netflix, Inc.': 'NFLX', 'Keurig Dr Pepper Inc.': 'KDP', 'Marqeta, Inc.': 'MQ', 'Sunrun Inc.': 'RUN', 'Johnson & Johnson': 'JNJ', 'Full Truck Alliance Co. Ltd.': 'YMM', 'Roku, Inc.': 'ROKU', 'Activision Blizzard, Inc.': 'ATVI', 'Array Technologies, Inc.': 'ARRY', 'B2Gold Corp.': 'BTG', 'RLX Technology Inc.': 'RLX', 'Bloom Energy Corporation': 'BE', 'United Microelectronics Corporation': 'UMC', 'ON Semiconductor Corporation': 'ON', 'Globus Medical, Inc.': 'GMED', 'Walgreens Boots Alliance, Inc.': 'WBA', 'JPMorgan Chase & Co.': 'JPM', 'The Williams Companies, Inc.': 'WMB', 'Range Resources Corporation': 'RRC', 'CNH Industrial N.V.': 'CNHI', 'Enphase Energy, Inc.': 'ENPH', 'Crescent Point Energy Corp.': 'CPG', 'Alteryx, Inc.': 'AYX', 'DexCom, Inc.': 'DXCM', 'Coty Inc.': 'COTY', 'Starbucks Corporation': 'SBUX', 'Microchip Technology Incorporated': 'MCHP', 'Viatris Inc.': 'VTRS', 'Magna International Inc.': 'MGA', 'MGM Resorts International': 'MGM', 'PepsiCo, Inc.': 'PEP', 'United Airlines Holdings, Inc.': 'UAL', 'Cameco Corporation': 'CCJ', 'CEMEX, S.A.B. de C.V.': 'CX', 'Antero Resources Corporation': 'AR', 'Airbnb, Inc.': 'ABNB', 'Regions Financial Corporation': 'RF', 'Merck & Co., Inc.': 'MRK', 'Amcor plc': 'AMCR', 'JetBlue Airways Corporation': 'JBLU', 'Tencent Music Entertainment Group': 'TME', 'The Goodyear Tire & Rubber Company': 'GT', 'Baker Hughes Company': 'BKR', 'Yamana Gold Inc.': 'AUY', 'Permian Resources Corporation': 'PR', 'Altria Group, Inc.': 'MO', 'Genworth Financial, Inc.': 'GNW', 'Wayfair Inc.': 'W', 'Franklin Resources, Inc.': 'BEN', 'Fortinet, Inc.': 'FTNT', 'AstraZeneca PLC': 'AZN', 'Telefonaktiebolaget LM Ericsson (publ)': 'ERIC', 'QuantumScape Corporation': 'QS', 'Gilead Sciences, Inc.': 'GILD', 'Mondelez International, Inc.': 'MDLZ', 'BlackBerry Limited': 'BB', 'International Flavors & Fragrances Inc.': 'IFF', 'Texas Instruments Incorporated': 'TXN', 'Boston Scientific Corporation': 'BSX', 'Cenovus Energy Inc.': 'CVE', 'Enterprise Products Partners L.P.': 'EPD', 'Invesco Ltd.': 'IVZ', 'DoorDash, Inc.': 'DASH', 'U.S. Bancorp': 'USB', 'Truist Financial Corporation': 'TFC', 'Peabody Energy Corporation': 'BTU', 'Raytheon Technologies Corporation': 'RTX', 'Oracle Corporation': 'ORCL', 'ZoomInfo Technologies Inc.': 'ZI', 'TAL Education Group': 'TAL', 'Morgan Stanley': 'MS', 'ArcelorMittal S.A.': 'MT', 'AppLovin Corporation': 'APP', 'Stellantis N.V.': 'STLA', 'Mattel, Inc.': 'MAT', 'International Business Machines Corporation': 'IBM', 'XP Inc.': 'XP', 'Patterson-UTI Energy, Inc.': 'PTEN', 'Dominion Energy, Inc.': 'D', 'Walmart Inc.': 'WMT'}

st.title("CFA Challenge Predictionapp")

st.markdown("<br>", unsafe_allow_html=True)

st.write('Monte Carlo simulations are valuable tools for this purpose. These simulations can predict the outcomes of processes or events influenced by random variables such as news'
         ', global events, and investor sentiment, all of which affect the stock market. This methodology was applied in my investment thesis during the CFA Research Challenge,'
          'demonstrating its effectiveness in financial market analysis.')

st.markdown("<br>", unsafe_allow_html=True)

st.write('Obtaining and displaying historical data next, we will generate historical financial data using yfinance. The desired interval of interest must be defined before we can pull the data. We will plot daily closing stock prices so we can focus on that information.')

st.markdown("<br>", unsafe_allow_html=True)

st.write("<h3 style='font-size:20px;'>Select a company</h3>", unsafe_allow_html=True)

name = st.selectbox('Select a company:', list(dict_name_tick.keys()))

ticker_name = dict_name_tick[name]

st.write(f'Ticker selected: {ticker_name}')

ticker = yf.Ticker(ticker_name)
start_date = '2010-01-01'
end_date = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
hist = ticker.history(start=start_date, end=end_date)

hist = hist[['Close']]

st.title('Stock prices - ' + name + ' (' + ticker_name + ')')
st.markdown("<span style='color:black; font-size:20px; font-weight:600'>Currency: $ </span>",unsafe_allow_html=True)
st.line_chart(hist)

# create day count, price and change lists

days = [i for i in range(1, len(hist['Close']) +1)]
price_orig = hist['Close'].to_list()
change = hist['Close'].pct_change().to_list()
# removing the first term since its NaN
change = change[1:]

# Statistics for use in Model
mean = np.mean(change)
std_dev = np.std(change)

st.markdown("<br>", unsafe_allow_html=True)

st.title('Statistics for stocks')

st.markdown("<span style='color:#31BBDE; font-size:20px; font-weight:700'>Mean percent change: </span>  {} %".format(str(round(mean*100,2))),unsafe_allow_html=True)
st.markdown("<span style='color:#31BBDE; font-size:20px; font-weight:700'>Standard Deviation of percent change: </span>  {} %".format(str(round(std_dev*100,2))),unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

st.title('Monte Carlo simulation')

st.write('To analyze our results, we need to set up a few things before we start the Monte Carlo simulation. In the first step, we need to specify how many simulations we would like to run, simulations, and how many days we would like to predict.')
st.write('A year of simulation (or 252 trading days) will be performed in this code (chosen arbitrarily, the more, the better). The original 11 years of data is also plotted, but it is cut after 2800 days to make it easier to visualize prediction lines.')
st.write(f'Our final step is track the year-out closing price of a prediction and if it’s higher than the {end_date} closing price.')

# Simulation Number and Prediction Period
st.write("<h3 style='font-size:20px;'>Input the number of simulations and number of trading simulation days:</h3>", unsafe_allow_html=True)
st.write('By default the simulation assumes 200 simulations and 252 trading days.')
simulations = st.number_input("Number of simulations:", value=200, min_value=1)
days_to_sim = st.number_input("Number of simulation days:", value=252, min_value=1)


# Initializing lists for analysis
close_end = []
above_close = []

# For loop for number of simulations desired
for i in range(simulations):
    num_days = [days[-1]]
    close_price = [hist.iloc[-1,0]]

    for j in range(days_to_sim):
        num_days.append(num_days[-1]+1)
        perc_change = norm.ppf(random(), loc = mean, scale=std_dev)
        close_price.append(close_price[-1]*(1+perc_change))

    if close_price[-1] > price_orig[-1]:
        above_close.append(1)
    else:
        above_close.append(0)

    close_end.append(close_price[-1])
    plt.plot(num_days, close_price)

plt.title(f'Monte Carlo stock prices - {simulations} simulations', fontweight='bold')
plt.xlabel('Trading Days after 01-01-2012', fontsize=12, fontweight='bold')
plt.ylabel('Closing Stock Price ($)', fontsize=12, fontweight='bold')

# Average Closing price and probability of increasing after 1 year

average_closing_price = sum(close_end) / simulations
average_perc_change = (average_closing_price-price_orig[-1])/price_orig[-1]

probability_of_increase = sum(above_close)/ simulations

# Adicionar linha tracejada com legenda
plt.axhline(y=average_closing_price, color='blue', linestyle='dashed', label='Average closing price')
plt.legend()

# Display the plot
st.pyplot(plt)

st.markdown("<br><br>", unsafe_allow_html=True)


# PREDICTIONS

st.title('Predictions after 1 year')

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("<span style='color:#31BBDE; font-size:20px; font-weight:700'>Predicted closing price after 1 year:  </span> $ {}".format(str(round(average_closing_price,2))),
            unsafe_allow_html=True)
st.markdown("<span style='color:#31BBDE; font-size:20px; font-weight:700'>Predicted percent increase after 1 year:  </span> {} %".format(str(round(average_perc_change*100,2))),
            unsafe_allow_html=True)
st.markdown("<span style='color:#31BBDE; font-size:20px; font-weight:700'>Probability of stock price increasing after 1 year:  </span> {} %".format(str(round(probability_of_increase*100,2))),
            unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

if average_perc_change<0:
    st.write(f'A year from now, the predicted average price will be ${str(round(average_closing_price,2))} ({str(round(average_perc_change*100,2))}%) higher than it was on {end_date}. The likelihood of {name} increasing after a year is also {str(round(probability_of_increase*100,2))}%.')
else:
    st.write(f'A year from now, the predicted average price will be ${str(round(average_closing_price, 2))} ({str(round(average_perc_change * 100, 2))}%) lower than it was on {end_date}. The likelihood of {name} increasing after a year is also {str(round((1-probability_of_increase) * 100, 2))}%.')

st.markdown("<br>", unsafe_allow_html=True)
st.write('Over a period of time, you will notice that the average price, average percent change, and probability of increasing are all convergent on a particular value. If you run the code with one simulation, the result will be randomized each time.')
