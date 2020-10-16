import os
import time
import datetime as dt
from configparser import ConfigParser
from pathlib import Path
import signal

import eikon as ek
import xlwings as xw

import warnings
warnings.filterwarnings("ignore")

# Eikon setup
conf = ConfigParser()
conf.read(os.path.join(os.path.dirname(__file__), '..', 'eikon.conf'))
print('Connecting to Eikon')
ek.set_app_key(conf['eikon']['APP_KEY'])


def main():
    pid_file = Path(__file__).resolve().parent / "pid"
    sheet = xw.Book.caller().sheets[0]
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
        with open(pid_file, 'w') as f:
            f.write(str(os.getpid()))
        sheet['C1'].value = 'running'

        sheet = xw.Book.caller().sheets[0]
        rng = sheet['A2'].expand()
        instruments, fields = rng[1:, 0].value, rng[0, 1:].value
        print('Connecting to streaming API')
        streaming_prices = ek.StreamingPrices(instruments=instruments, fields=fields)
        streaming_prices.open()
        print('Start pushing to Excel')

        while True:
            # Throttling to every half second
            rng[1, 1].value = streaming_prices.get_snapshot().values[:, 1:]
            sheet['B1'].value = dt.datetime.now()
            time.sleep(0.5)


if __name__ == '__main__':
    xw.Book('realtime.xlsx').set_mock_caller()
    main()
