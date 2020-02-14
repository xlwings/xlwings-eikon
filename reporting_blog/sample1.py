import eikon as ek
import xlwings as xw

# Please read your Eikon APP_KEY from a config file or environment variable
ek.set_app_key('YOUR_APP_KEY')

# Parameters
instrument = '.DJI'
start_date = '2020-01-01'
end_date = '2020-01-31'

# Request time series data from Eikon, will return a Pandas DataFrame
df = ek.get_timeseries(instrument,
                       fields='*',
                       start_date=start_date,
                       end_date=end_date)

# Open a new Excel workbook and write the Pandas DataFrame to A1 on the first sheet
wb = xw.Book()
wb.sheets[0]['A1'].value = df
