import os
import sys
import webbrowser
import threading
import time
import socket
import uvicorn

def is_port_in_use(port: int) -> bool:
    """Checks if a local TCP port is already open/listening."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            return s.connect_ex(('127.0.0.1', port)) == 0
    except Exception:
        return False

def open_browser(url: str = "http://127.0.0.1:8000"):
    time.sleep(1.0)
    print(f"\n[*] Opening web application interface at: {url}")
    webbrowser.open(url)

if __name__ == "__main__":
    print("================================================================")
    print("  BANKTECH VALUATION OCR & EXCEL REPORT GENERATOR (INDIA SHELTER)")
    print("================================================================")

    url = "http://127.0.0.1:8000"

    # If the server is already running (e.g. from share daemon or another window)
    if is_port_in_use(8000):
        print(f"\n[+] BankTech server is ALREADY running on port 8000!")
        print(f"[*] Launching your web browser to: {url}")
        open_browser(url)
        print("\n[*] Application is ready. You can close this terminal or press Ctrl+C.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            sys.exit(0)

    print(f"  Server starting on {url} ...")
    
    # Launch browser in background thread once server starts
    threading.Thread(target=open_browser, daemon=True).start()
    
    uvicorn.run("backend.app:app", host="0.0.0.0", port=8000, reload=False)
