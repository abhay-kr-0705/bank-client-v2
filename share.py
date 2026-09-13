"""
BankTech Live Client Preview Launcher via Cloudflare Tunnel
Launches the FastAPI application locally and creates an encrypted Cloudflare Tunnel.
Generates an instant, public HTTPS link (https://*.trycloudflare.com) that your client
can access remotely from anywhere, running directly from your PC.
"""

import os
import sys
import time
import re
import subprocess
import threading
import webbrowser

# Ensure unbuffered output so status prints immediately in all terminals
_print = print
def print(*args, **kwargs):
    kwargs["flush"] = True
    _print(*args, **kwargs)

# Safe UTF-8 console encoding on Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

CLOUDFLARED_CANDIDATE_PATHS = [
    r"D:\DOWNLOADS\cloudflared-windows-amd64.exe",
    r"C:\Program Files\Cloudflare\cloudflared.exe",
    r"C:\Program Files (x86)\Cloudflare\cloudflared.exe",
    r"C:\Cloudflare\cloudflared.exe",
    "cloudflared"
]

def find_cloudflared() -> str:
    """Finds the cloudflared executable path on this machine."""
    for p in CLOUDFLARED_CANDIDATE_PATHS:
        if os.path.isfile(p):
            return p
    # Check if in system PATH
    try:
        res = subprocess.run(["where.exe", "cloudflared"], capture_output=True, text=True)
        if res.returncode == 0 and res.stdout.strip():
            return res.stdout.strip().splitlines()[0]
    except Exception:
        pass
    return r"D:\DOWNLOADS\cloudflared-windows-amd64.exe"

def copy_to_clipboard(text: str):
    """Copies the live tunnel URL to the Windows clipboard."""
    try:
        subprocess.run("clip", input=text.strip().encode("utf-8"), check=False, shell=True)
    except Exception:
        pass

def run_uvicorn_in_thread():
    """Starts the FastAPI app in-process on port 8000."""
    import uvicorn
    uvicorn.run("backend.app:app", host="127.0.0.1", port=8000, log_level="warning")

def run():
    print("=" * 72)
    print("   BANKTECH VALUATION OCR & EXCEL ENGINE -- LIVE CLIENT SHARING")
    print("=" * 72)

    cf_bin = find_cloudflared()
    if not os.path.isfile(cf_bin) and cf_bin != "cloudflared":
        print(f"[!] Error: cloudflared not found at {cf_bin}")
        print("    Please ensure cloudflared is installed or present on your PC.")
        sys.exit(1)

    print(f"[*] Found Cloudflare binary: {cf_bin}")
    print("[*] Starting local web application server on port 8000...")
    
    server_thread = threading.Thread(target=run_uvicorn_in_thread, daemon=True)
    server_thread.start()

    # Wait for server to bind
    time.sleep(2)

    print("[*] Launching Cloudflare Tunnel (connecting to local app)...")
    cf_cmd = [cf_bin, "tunnel", "--url", "http://127.0.0.1:8000"]
    cf_proc = subprocess.Popen(
        cf_cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        encoding="utf-8",
        errors="replace"
    )

    tunnel_url = None
    # Matches quick tunnel generated subdomains with hyphens e.g. https://slope-yes-corporation-optics.trycloudflare.com
    url_pattern = re.compile(r"https://([a-zA-Z0-9]+-[a-zA-Z0-9\-]+)\.trycloudflare\.com")

    try:
        for line in cf_proc.stdout:
            match = url_pattern.search(line)
            if match and not tunnel_url:
                tunnel_url = f"https://{match.group(1)}.trycloudflare.com"
                copy_to_clipboard(tunnel_url)
                print("\n" + "#" * 72)
                print("  [+] LIVE CLIENT PREVIEW URL GENERATED:")
                print(f"      {tunnel_url}")
                print("#" * 72)
                print("  [*] URL copied to your clipboard!")
                print("  [*] Anyone with this link can access and use your app remotely.")
                print("  [*] Traffic is securely proxied to your local PC via Cloudflare.")
                print("-" * 72)
                print("  Local URL:  http://127.0.0.1:8000")
                print("  Remote URL: " + tunnel_url)
                print("-" * 72)
                print("  Press Ctrl + C at any time to stop sharing.\n")
                
                # Open browser
                webbrowser.open(tunnel_url)

    except KeyboardInterrupt:
        print("\n[*] Stopping live share...")
    finally:
        try:
            cf_proc.terminate()
            cf_proc.kill()
        except Exception:
            pass
        print("[*] Live preview session closed safely.")

if __name__ == "__main__":
    run()
