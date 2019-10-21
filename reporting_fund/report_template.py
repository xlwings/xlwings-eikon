import os
import pickle
import datetime as dt
import xlwings as xw
from xlwings_reports import create_report  # not part of the open-source xlwings package
import numpy as np
import pandas as pd
import eikon as ek
from configparser import ConfigParser


def main():
    # Files
    template = xw.Book.caller()
    template_path = template.fullname
    report_path = os.path.join(os.path.dirname(template_path), 'fund_report.xlsx')

    # Eikon
    conf = ConfigParser()
    conf.read(os.path.join(os.path.dirname(template_path), '..', 'eikon.conf'))
    ek.set_app_key(conf['eikon']['APP_KEY'])

    # Configuration
    date_format = template.sheets['Config']['date_format'].value
    if date_format == 'UK':
        fmt = '%e %b %Y'
    elif date_format == 'US':
        fmt = '%b %e, %Y'
    else:
        fmt = '%e %b %Y'

    instrument = template.sheets['Config']['instrument'].value

    # Get your data from Eikon
    now = dt.datetime.now()
    perf_start_date = dt.datetime(now.year - 4, 1, 1)

    historical_perf = ek.get_timeseries(instrument,
                             fields=['close'],
                             start_date=perf_start_date,
                             end_date=now,
                             interval='weekly')

    perf_end_date = historical_perf.index[-1]

    ret, err = ek.get_data(instrument, fields=['TR.IndexName', 'TR.IndexCalculationCurrency'])
    index_name = ret.loc[0, 'Index Name']
    currency = ret.loc[0, 'Calculation Currency']

    constituents, err = ek.get_data(f'0#{instrument}', fields=['TR.CommonName', 'TR.PriceClose', 'TR.TotalReturnYTD'])

    constituents = constituents.set_index('Company Common Name')
    for i in range(0, 6):
        constituents.insert(loc=i, column='merged' + str(i), value=np.nan)
    constituents = constituents.drop(['Instrument'], axis=1)

    constituents = constituents.rename(columns={"YTD Total Return": "YTD %"})

    # Summary
    summary, err = ek.get_data(instrument, ['TR.PriceClose', 'TR.Volume', 'TR.PriceLow', 'TR.PriceHigh'])



    # Collect all data
    data = dict(
        perf_start_date=perf_start_date.strftime(fmt),
        perf_end_date=perf_end_date.strftime(fmt),
        index_name=index_name,
        currency=currency,
        reference_date=dt.date.today().strftime(fmt),
        historical_perf=historical_perf,
        constituents=constituents,
        price_close=float(summary['Price Close']),
        volume=float(summary['Volume']),
        price_low=float(summary['Price Low']),
        price_high=float(summary['Price High'])
    )

    # Create the Excel report
    wb = create_report(template_path, report_path, **data)


if __name__ == '__main__':
    # This part is to run the script directly from Python, not via Excel
    xw.Book(os.path.join(os.path.dirname(__file__), 'fund_template.xlsx')).set_mock_caller()
    main()
