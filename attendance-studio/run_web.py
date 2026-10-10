"""Run Attendance Studio in the default browser on this computer."""
import threading
import os
import webbrowser
import time
import json
import socket
from urllib.request import urlopen
import uvicorn


def attendance_running(port):
    try:
        with urlopen(f'http://127.0.0.1:{port}/api/health', timeout=1) as response:
            health = json.load(response)
        with urlopen(f'http://127.0.0.1:{port}/openapi.json', timeout=1) as response:
            schema = json.load(response)
        return health.get('status') == 'ok' and schema.get('info', {}).get('title') == 'Attendance Studio'
    except (OSError, ValueError, AttributeError):
        return False


def open_when_ready(port):
    for _ in range(60):
        if attendance_running(port):
            webbrowser.open(f'http://127.0.0.1:{port}/#/login')
            return
        time.sleep(.5)


def main():
    host = os.getenv('HOST', '127.0.0.1')
    port = int(os.getenv('PORT', '8877'))
    open_browser = os.getenv('OPEN_BROWSER', '1') == '1'
    if attendance_running(port):
        print(f'Attendance Studio is already running: http://127.0.0.1:{port}/#/login')
        if open_browser:
            webbrowser.open(f'http://127.0.0.1:{port}/#/login')
        return
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
        try:
            listener.bind((host, port))
        except OSError as error:
            print(f'Cannot start Attendance Studio on {host}:{port}. The port may be in use. {error}')
            raise SystemExit(1)
        listener.listen(128)
        if open_browser:
            threading.Thread(target=open_when_ready, args=(port,), daemon=True).start()
        uvicorn.Server(uvicorn.Config('backend.main:app', host=host, port=port, use_colors=False)).run(sockets=[listener])


if __name__ == '__main__':
    main()
