import os
import datetime as dt
from configparser import ConfigParser
from pathlib import Path
import signal

import refinitiv.dataplatform as rdp
import xlwings as xw

rng = None  # global


def send_snapshot_to_excel(streaming_prices):
    rng[1, 1].value = streaming_prices.get_snapshot().to_numpy()[:, 1:]
    rng[0, 0].offset(row_offset=-1, column_offset=1).value = dt.datetime.now()


def main():
    # Excel
    global rng
    sheet = xw.Book.caller().sheets[0]
    rng = sheet['A2'].expand()

    pid_file = Path(__file__).resolve().parent / "pid"
    if pid_file.exists():
        # Stop streaming
        with open(pid_file, 'r') as f:
            pid = f.read()
        try:
            os.kill(int(pid), signal.SIGSTOP)
            pid_file.unlink()
        except ProcessLookupError as e:
            os.remove(pid_file)
        finally:
            rng[1:, 1:].clear_contents()
            rng[0, 0].offset(row_offset=-1, column_offset=1).clear_contents()
    else:  # Start streaming
        with open(pid_file, 'w') as f:
            f.write(str(os.getpid()))

        # Connect to Refinitiv Workspace
        conf = ConfigParser()
        conf.read(os.path.join(os.path.dirname(__file__), '..', 'eikon.conf'))
        rdp.open_desktop_session(conf['eikon']['APP_KEY'])

        instruments, fields = rng[1:, 0].value, rng[0, 1:].value
        streaming_prices = rdp.StreamingPrices(
            universe=instruments,
            fields=fields,
            on_update=lambda streaming_price, instrument_name, fields:
            send_snapshot_to_excel(streaming_price)
        )

        while True:
            streaming_prices.open()


if __name__ == '__main__':
    xw.Book('realtime_eikon.xlsx').set_mock_caller()
    main()
