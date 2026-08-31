import os
import sys
import webbrowser
import threading
import time
import uvicorn

def open_browser():
    time.sleep(1.2)
    url = "http://127.0.0.1:8000"
    print(f"\n[*] Opening web application interface at: {url}")
    webbrowser.open(url)

if __name__ == "__main__":
    print("================================================================")
    print("  BANKTECH VALUATION OCR & EXCEL REPORT GENERATOR (INDIA SHELTER)")
    print("================================================================")
    print("  Server starting on http://127.0.0.1:8000 ...")
    
    # Launch browser in background thread
    threading.Thread(target=open_browser, daemon=True).start()
    
    uvicorn.run("backend.app:app", host="127.0.0.1", port=8000, reload=False)
