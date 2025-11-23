"""
Diagnostic script to check internet connectivity for voice recognition.
This script tests the same internet check that NetworkMonitor uses.
"""
import socket
import sys

def test_internet_connection():
    """Test connection to Google DNS (8.8.8.8:53) the same way NetworkMonitor does."""
    print("Testing internet connectivity...")
    print("=" * 60)
    
    # Test 1: Connect to Google DNS (8.8.8.8:53)
    print("\nTest 1: Connecting to Google DNS (8.8.8.8:53)...")
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        print("[SUCCESS] Connected to 8.8.8.8:53")
        print("   Google Voice-to-Text should work!")
        dns_works = True
    except OSError as e:
        print(f"[FAILED] {e}")
        print("   This is why Vosk is being used instead of Google!")
        dns_works = False
    
    # Test 2: Alternative DNS server (1.1.1.1:53 - Cloudflare)
    print("\nTest 2: Connecting to Cloudflare DNS (1.1.1.1:53)...")
    try:
        socket.create_connection(("1.1.1.1", 53), timeout=3)
        print("[SUCCESS] Connected to 1.1.1.1:53")
        cloudflare_works = True
    except OSError as e:
        print(f"[FAILED] {e}")
        cloudflare_works = False
    
    # Test 3: HTTP connection to google.com
    print("\nTest 3: Connecting to google.com:80...")
    try:
        socket.create_connection(("google.com", 80), timeout=3)
        print("[SUCCESS] Connected to google.com:80")
        http_works = True
    except OSError as e:
        print(f"[FAILED] {e}")
        http_works = False
    
    # Summary
    print("\n" + "=" * 60)
    print("DIAGNOSIS SUMMARY:")
    print("=" * 60)
    
    if dns_works:
        print("[OK] Your internet is working correctly!")
        print("   The system SHOULD use Google Voice-to-Text.")
        print("\n[WARNING] If Vosk is still being used, there may be another issue:")
        print("   - Check for firewall blocking during app startup")
        print("   - Try running as administrator")
        print("   - Check the server logs for the exact error")
    elif cloudflare_works or http_works:
        print("[WARNING] Internet works, but Google DNS (8.8.8.8) is blocked!")
        print("   Possible causes:")
        print("   - Corporate/School firewall blocking DNS port 53")
        print("   - VPN or antivirus blocking specific DNS servers")
        print("   - ISP blocking Google DNS")
        print("\n[SOLUTION] The code can be modified to use alternative checks")
    else:
        print("[ERROR] No internet connectivity detected!")
        print("   This is why the system is using Vosk offline mode.")
        print("   Check your network connection.")
    
    print("=" * 60)
    return dns_works

if __name__ == "__main__":
    result = test_internet_connection()
    sys.exit(0 if result else 1)
