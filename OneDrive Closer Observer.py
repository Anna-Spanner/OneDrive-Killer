import time, threading, psutil, subprocess, wmi, os, ctypes
from http.server import SimpleHTTPRequestHandler, HTTPServer
status = False
program = "OneDrive Closer.exe"
win = wmi.WMI()

def checkCloserOpen():
    try:
        os.startfile(program)
    except:
        ctypes.windll.user32.MessageBoxW(0, f"Something went wrong when opening {program}! Exiting OneDrive Closer, please re-launch or contact Anna-Rose. Make sure to say 'error caused in first time'", "Error", 0x10)
        time.sleep(1)
        exit()
    time.sleep(5)
    if not findAppOpen(program):
        print(program, "still isn't open... trying again to open")
        os.startfile(program)
        time.sleep(5)
        if not findAppOpen(program):
            print(program, "wasn't found open again... last try")
            os.startfile(program)
            time.sleep(5)
            if not findAppOpen(program):
                ctypes.windll.user32.MessageBoxW(0, f"Something went wrong when opening {program}! Exiting OneDrive Closer, please re-launch or contact Anna-Rose.", "Error", 0x10)
                time.sleep(1)
                exit()

def findAppOpen(app: str):
    confirm = False
    for process in win.Win32_Process():
        if process.Name == app:
            confirm = True
            break
    return confirm

def close(application):
    for proc in psutil.process_iter(['pid', 'name', 'username']):
        if proc.info['name'] == application:
            try:
                if proc.info['username'] == 'SYSTEM' or proc.info['username'] is None:
                    proc.terminate()
                else:
                    proc.terminate()
            except psutil.AccessDenied:
                print("Normal way failed, killing through cmd")
                subprocess.run(['taskkill', '/F', '/PID', str(proc.info['pid'])], check=True)

def set_status():
    global status
    status = True

def handle_request():
    global status
    class RequestHandler(SimpleHTTPRequestHandler):
        def do_GET(self):
            if self.path == "/ping":
                set_status()
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b'Ping received! Status set to True.')
    server = HTTPServer(('localhost', 46237), RequestHandler)
    print("Server is running on localhost:46237...")
    server.serve_forever()

def check_status():
    global status
    while True:
        time.sleep(65)
        if status:
            status = False 
        else:
            print("No update from OneDrive Closer...")
            time.sleep(60)
            if status:
                status = False
            else:
                print("Checking if OneDrive Closer is open...")
                if findAppOpen(program):
                    close(program)
                checkCloserOpen()


if __name__ == "__main__":
    server_thread = threading.Thread(target=handle_request, daemon=True)
    server_thread.start()
    if not findAppOpen(program):
        checkCloserOpen()

    check_status()
