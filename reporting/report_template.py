import os
import datetime as dt
from configparser import ConfigParser

import numpy as np
import eikon as ek
import xlwings as xw
from xlwings_reports import create_report  # not part of the open-source xlwings package


def main():
    # Files
    template = xw.Book.caller()
    report_path = os.path.join(os.path.dirname(__file__), 'report.xlsx')

    # Eikon setup
    conf = ConfigParser()
    conf.read(os.path.join(os.path.dirname(__file__), '..', 'eikon.conf'))
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
    start_date = dt.datetime(dt.datetime.now().year - 4, 1, 1)

    # Prices
    prices = ek.get_timeseries(instrument,
                               fields=['close'],
                               start_date=start_date,
                               end_date=dt.datetime.now(),
                               interval='weekly')
    end_date = prices.index[-1]

    # Summary
    summary, err = ek.get_data(instrument,
                               ['TR.PriceClose', 'TR.Volume', 'TR.PriceLow', 'TR.PriceHigh',
                                'TR.IndexName', 'TR.IndexCalculationCurrency'])

    # Constituents
    constituents, err = ek.get_data(f'0#{instrument}',
                                    fields=['TR.CommonName', 'TR.PriceClose', 'TR.TotalReturnYTD'])
    constituents = constituents.set_index('Company Common Name')
    # Add empty columns so it goes into the desired Excel cells
    for i in range(0, 6):
        constituents.insert(loc=i, column='merged' + str(i), value=np.nan)
    constituents = constituents.drop(['Instrument'], axis=1)
    constituents = constituents.rename(columns={"YTD Total Return": "YTD %"})

    # Collect data
    data = dict(
        perf_start_date=start_date.strftime(fmt),
        perf_end_date=end_date.strftime(fmt),
        index_name=summary.loc[0, 'Index Name'],
        currency=summary.loc[0, 'Calculation Currency'],
        reference_date=dt.date.today().strftime(fmt),
        historical_perf=prices,
        constituents=constituents,
        price_close=float(summary['Price Close']),
        volume=float(summary['Volume']),
        price_low=float(summary['Price Low']),
        price_high=float(summary['Price High'])
    )

    # Create the Excel report
    wb = create_report(template.fullname, report_path, **data)


if __name__ == '__main__':
    # This part is to run the script directly from Python, not via Excel
    xw.books.active.set_mock_caller()
    main()
