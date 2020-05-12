import os
import eikon as ek
import xlwings as xw
# Requires a license key: https://www.xlwings.org/trial
from xlwings.pro.reports import create_report

# Please read your Eikon APP_KEY from a config file or environment variable
ek.set_app_key('YOUR_APP_KEY')

def main():
    # These parameters could also come from a config sheet in the template
    instrument = '.DJI'
    start_date = '2020-01-01'
    end_date = '2020-01-31'

    # Eikon queries
    df = ek.get_timeseries(instrument,
                           fields='*',
                           start_date=start_date,
                           end_date=end_date)

    summary, err = ek.get_data(instrument,
                               ['TR.IndexName', 'TR.IndexCalculationCurrency'])

    # Populate the Excel template with the data
    template = xw.Book.caller().fullname
    wb = create_report(template=template,
                       output=os.path.join(os.path.dirname(template), 'factsheet.xlsx'),
                       title=f"{summary.loc[0, 'Index Name']} ({summary.loc[0, 'Calculation Currency']})",
                       df=df)


if __name__ == '__main__':
    xw.Book('sample3.xlsx').set_mock_caller()
    main()