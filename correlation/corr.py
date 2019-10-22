import os
from configparser import ConfigParser
import datetime as dt
from dateutil.relativedelta import relativedelta

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import eikon as ek
import xlwings as xw


conf = ConfigParser()
conf.read(os.path.join(os.path.dirname(__file__), '..', 'eikon.conf'))
ek.set_app_key(conf['eikon']['APP_KEY'])


@xw.func
# @xw.ret(expand='table')  # use this if your version of Excel doesn't have dynamic arrays
def eikon_get_corr(rics, start_date=None, end_date=None, fields=None, interval='daily'):
    if fields is None:
        fields = ['close']
    if start_date is None:
        start_date = dt.datetime.now() - relativedelta(years=1)
    if end_date is None:
        end_date = dt.datetime.now()
    # Eikon query
    prices = ek.get_timeseries(rics, fields=fields,
                               start_date=start_date,
                               end_date=end_date,
                               interval=interval)
    # Clean data, calculate log returns and correlation
    prices = prices.dropna()
    log_ret = np.log(prices / prices.shift(1))
    corr = log_ret.corr()
    corr.index.name = None
    return corr


@xw.func
@xw.arg('corr', pd.DataFrame)
def corr_plot(corr):
    wb = xw.Book.caller()
    # Seaborn heatmap
    ax = sns.heatmap(corr, cmap='coolwarm', vmin=-1, vmax=1, linewidths=.5,
                     xticklabels=True, yticklabels=True)
    ax.tick_params(left=False, bottom=False)
    plt.yticks(rotation=0)
    plt.xticks(rotation=90)
    # Pass to Excel as picture
    fig = ax.get_figure()
    wb.sheets.active.pictures.add(fig,
                                  top=wb.selection.top,
                                  left=wb.selection.left,
                                  name='CorrPlot',
                                  update=True)
    plt.close()
    return '<Corr Plot>'
