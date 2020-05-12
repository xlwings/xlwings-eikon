import os
import math
from dateutil.relativedelta import relativedelta

import numpy as np
import pandas as pd
import eikon as ek
import xlwings as xw


def main():
    # Eikon
    ek.set_app_key(os.getenv('EIKON_APP_KEY'))

    # Excel
    sheet = xw.Book.caller().sheets[0]
    sheet['O1'].expand().clear_contents()
    instrument = sheet['E5'].value
    start_date = sheet['E4'].value
    end_date = start_date + relativedelta(years=2)

    # History
    prices = ek.get_timeseries(instrument,
                               fields='close',
                               start_date=start_date,
                               end_date=end_date)

    # Take annualized mean and standard deviation from first 252 trading days
    trading_days = 252
    if len(prices) < trading_days:
        raise Exception('History must have at least 252 data points.')
    returns = np.log(prices[:trading_days] / prices[:trading_days].shift(1))
    mean = float(np.mean(returns) * trading_days)
    stdev = float(returns.std() * math.sqrt(trading_days))

    # Simulation parameters
    num_simulations = sheet['E3'].options(numbers=int).value
    time = 1  # years
    num_timesteps = len(pd.date_range(prices[:trading_days].index[-1], end_date, freq='B'))
    dt = time/num_timesteps  # Length of time period
    vol = stdev
    mu = mean  # Drift
    starting_price = float(prices.iloc[trading_days - 1])
    percentile_selection = [5, 50, 95]

    # Preallocation and intial values
    price = np.zeros((num_timesteps, num_simulations))
    percentiles = np.zeros((num_timesteps, 3))
    price[0, :] = starting_price
    percentiles[0, :] = starting_price

    # Simulation at each time step (log normal distribution)
    for t in range(1, num_timesteps):
        rand_nums = np.random.randn(num_simulations)
        price[t, :] = price[t - 1, :] * np.exp((mu - 0.5 * vol**2) * dt + vol * rand_nums * np.sqrt(dt))
        percentiles[t, :] = np.percentile(price[t, :], percentile_selection)

    # Turn into pandas DataFrame
    simulation = pd.DataFrame(data=percentiles,
                              index=pd.date_range(prices[:trading_days].index[-1], end_date, freq='B'),
                              columns=['5th Percentile', 'Median', '95th Percentile'])

    # Concat history & simulation and reorder cols
    combined = pd.concat([prices, simulation], axis=1)
    combined = combined[['5th Percentile', 'Median', '95th Percentile', 'CLOSE']]
    sheet['O1'].value = combined
    sheet.charts['Chart 3'].set_source_data(sheet['O1'].expand())


if __name__ == '__main__':
    # This part is to run the script directly from Python, not via Excel
    xw.Book("simulation.py").set_mock_caller()
    main()
