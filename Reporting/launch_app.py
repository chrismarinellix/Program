# robust_launch.py  (or launch_app.py)

"""
Double-click or run `python robust_launch.py`
to start Streamlit, wait until it’s live, and open the browser.
Detailed logs go to logs/streamlit_YYYYmmdd_HHMMSS.log
"""

import subprocess, sys, time, webbrowser, os, pathlib, socket, datetime, threading

###############################################################################
APP_PATH = pathlib.Path(r"C:\Reporting\Py\Program\report_streamlit.py")  # ← your app
PORT     = 8501
HEADLESS = "false"        # "false" = not headless, Streamlit will open to LAN
LOG_DIR  = pathlib.Path("logs")
LOG_DIR.mkdir(exist_ok=True)
###############################################################################

log_file = LOG_DIR / f"streamlit_{datetime.datetime.now():%Y%m%d_%H%M%S}.log"
LOG_FH   = open(log_file, "w", encoding="utf-8")

def log(msg: str):
    print(msg)
    LOG_FH.write(msg + "\n")
    LOG_FH.flush()

def port_in_use(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("localhost", port)) == 0

# ------------------------------------------------------------------ sanity check
if not APP_PATH.exists():
    log(f"🚫  Streamlit app not found: {APP_PATH}")
    sys.exit(1)

if port_in_use(PORT):
    log(f"🚫  Port {PORT} is already in use. "
        "Is another Streamlit instance running?")
    sys.exit(1)

# ------------------------------------------------------------------ launch
cmd = [sys.executable, "-m", "streamlit", "run", str(APP_PATH),
       "--server.headless", HEADLESS]

log("▶️  Launching Streamlit:\n    " + " ".join(cmd))
proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT, text=True)

# forward child output to console + log
def pump():
    for line in proc.stdout:
        LOG_FH.write(line)
        LOG_FH.flush()
        sys.stdout.write(line)
threading.Thread(target=pump, daemon=True).start()

# wait up to 15 s for the server to bind the port
timeout = 15
start = time.perf_counter()
while time.perf_counter() - start < timeout:
    if port_in_use(PORT):
        url = f"http://localhost:{PORT}"
        log(f"✅  Streamlit is live – opening {url}")
        webbrowser.open_new_tab(url)
        break
    if proc.poll() is not None:
        log("❌  Streamlit process terminated early – check log above.")
        LOG_FH.close()
        sys.exit(proc.returncode)
    time.sleep(0.4)
else:
    log(f"❌  Gave up waiting after {timeout} s – see log file for details.")
    proc.terminate()
    LOG_FH.close()
    sys.exit(1)

# keep launcher alive, forward CTRL-C
try:
    while proc.poll() is None:
        time.sleep(0.5)
except KeyboardInterrupt:
    log("⏹  CTRL-C – terminating Streamlit …")
    proc.terminate()
    proc.wait()

LOG_FH.close()
