import wmi, time, os, ctypes, requests, psutil, subprocess

win = wmi.WMI()
observer = "OneDrive Closer Observer.exe"
programToClose = "OneDrive.exe"

def findAppOpen(app: str):
    confirm = False
    for process in win.Win32_Process():
        if process.Name == app:
            confirm = True
            break
    return confirm

def checkObserverOpen():
    if not findAppOpen(observer):
        print(observer, "wasn't found open... attempting to open")
        os.startfile(observer)
        time.sleep(5)
        if not findAppOpen(observer):
            print(observer, "still isn't open... trying again to open")
            os.startfile(observer)
            time.sleep(5)
            if not findAppOpen(observer):
                print(observer, "wasn't found open again... last try")
                os.startfile(observer)
                time.sleep(5)
                if not findAppOpen(observer):
                    ctypes.windll.user32.MessageBoxW(0, f"Something went wrong when opening {observer}! Exiting OneDrive Closer, please re-launch or contact Anna-Rose.", "Error", 0x10)
                    time.sleep(1)
                    exit()

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

def observerChecker():
    try:
        response = requests.get("http://localhost:46237/ping")
        if response.status_code == 200:
            return True
        else:
            return False
    except Exception as e:
        print(e)
        return False

if __name__ == "__main__":
    while True:
        try:
            if not observerChecker():
                print("No response from observer, checking if its open")
                if findAppOpen(observer):
                    print("Observer still open, closing then restarting.")
                    close(observer)
                    time.sleep(2)
                    checkObserverOpen()
                else:
                    checkObserverOpen()
            print(f"Checking if {programToClose} is open")
        except:
            ctypes.windll.user32.MessageBoxW(0, f"Something went wrong during the observer check! Exiting OneDrive Closer, please re-launch. If the issue persists ensure or this application and {observer} are closed before relaunching. Still having issues? Contact Anna-Rose.", "Observer Error", 0x10)
        try:
            if findAppOpen(programToClose):
                print("Found! Closing")
                close(programToClose)
        except:
            ctypes.windll.user32.MessageBoxW(0, f"Something went wrong when closing {programToClose}! Exiting OneDrive Closer, please re-launch. If the issue persists ensure or this application and {observer} are closed before relaunching. Still having issues? Contact Anna-Rose.", "Closing Error", 0x10)
        print("sleeping...")
        time.sleep(60)
