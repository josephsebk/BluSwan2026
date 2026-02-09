import os
import requests
import json
import glob
import re

# Configuration
ASSETS_DIR = 'assets/investors'
INDEX_HTML = 'index.html'

def get_apps_script_url():
    """Extracts the Apps Script URL from index.html"""
    try:
        with open(INDEX_HTML, 'r') as f:
            content = f.read()
            match = re.search(r'const APPS_SCRIPT_URL = "(https://script.google.com/macros/s/[^"]+/exec)"', content)
            if match:
                return match.group(1)
    except Exception as e:
        print(f"Error reading {INDEX_HTML}: {e}")
    return None

def reset_local_files():
    """Deletes all .webp files in assets/investors/"""
    print(f"Cleaning {ASSETS_DIR}...")
    files = glob.glob(os.path.join(ASSETS_DIR, '*.webp'))
    for f in files:
        try:
            os.remove(f)
            print(f"Deleted {f}")
        except Exception as e:
            print(f"Error deleting {f}: {e}")
    
    # Also remove the mapping file
    mapping_file = 'assets/investor_images.js'
    if os.path.exists(mapping_file):
        try:
            os.remove(mapping_file)
            print(f"Deleted {mapping_file}")
        except Exception as e:
             print(f"Error deleting {mapping_file}: {e}")

def reset_google_sheet(url):
    """Resets all started meetings to pending via the GAS endpoint"""
    print(f"Connecting to Google Sheet via {url}...")
    
    try:
        # 1. Get current meetings
        response = requests.get(url, params={'action': 'getMeetings'})
        if response.status_code != 200:
            print(f"Failed to fetch meetings: {response.text}")
            return

        data = response.json()
        meetings = data.get('meetings', [])
        
        started_meetings = [m for m in meetings if m.get('status', '').lower() == 'started']
        
        if not started_meetings:
            print("No started meetings found to reset.")
            return

        print(f"Found {len(started_meetings)} started meetings. Resetting...")

        # 2. Reset each started meeting
        headers = {'Content-Type': 'text/plain'} # GAS often prefers plain text for post body or simple JSON string
        for m in started_meetings:
            payload = json.dumps({
                'action': 'resetMeeting',
                'meetingId': m['id']
            })
            
            # Note: Requests to GAS macros sometimes follow redirects.
            res = requests.post(url, data=payload, headers=headers)
            
            if res.status_code == 200:
                print(f"Reset meeting {m['id']} ({m.get('founder', 'Unknown')}): OK")
            else:
                print(f"Failed to reset meeting {m['id']}: {res.status_code} {res.text}")

    except Exception as e:
        print(f"Error resetting sheet: {e}")

def main():
    print("=== BluSwan App Reset Tool ===")
    
    # 1. Local Reset
    reset_local_files()
    
    # 2. Remote Reset
    url = get_apps_script_url()
    if url and "script.google.com" in url:
        reset_google_sheet(url)
    else:
        print("Could not find valid APPS_SCRIPT_URL in index.html. Skipping sheet reset.")
    
    print("=== Reset Complete ===")

if __name__ == "__main__":
    main()
