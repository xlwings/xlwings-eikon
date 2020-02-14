import eikon as ek
import xlwings as xw

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
    wb = xw.Book.caller()
    wb.sheets[0]['A4'].value = f"{summary.loc[0, 'Index Name']} ({summary.loc[0, 'Calculation Currency']})"
    wb.sheets[0]['A6'].value = df

if __name__ == '__main__':
    xw.Book('sample2.xlsx').set_mock_caller()
    main()