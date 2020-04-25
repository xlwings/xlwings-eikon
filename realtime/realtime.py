import os
import signal
import time
from pathlib import Path
import datetime as dt
from configparser import ConfigParser

import pandas as pd
import eikon as ek
import xlwings as xw

# Eikon setup
conf = ConfigParser()
conf.read(os.path.join(os.path.dirname(__file__), '..', 'eikon.conf'))
ek.set_app_key(conf['eikon']['APP_KEY'])


def main():
    sheet = xw.Book.caller().sheets[0]
    rng = sheet['A2'].expand()
    tickers, fields = rng[1:, 0].value, rng[0, 1:].value
    df, err = ek.get_data(tickers, fields)

    pid_file = Path(__file__).resolve().parent / "pid"
    if pid_file.exists():
        # Stop server
        with open(pid_file, 'r') as f:
            pid = f.read()
        try:
            os.kill(int(pid), signal.SIGSTOP)
            os.remove(pid_file)
        except ProcessLookupError as e:
            os.remove(pid_file)
        finally:
            sheet['C1'].value = 'stopped'
    else:
        # Start server
        with open(pid_file, 'w') as f:
            f.write(str(os.getpid()))
        sheet['C1'].value = 'running'

        while True:
            if not df.equals(rng.options(pd.DataFrame, index=False).value):
                rng[1, 1].value = df.values[:, 1:]
                sheet['B1'].value = dt.datetime.now()
            time.sleep(2)


if __name__ == '__main__':
    xw.books.active.set_mock_caller()
    main()
