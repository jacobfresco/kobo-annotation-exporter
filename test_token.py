import json
import requests
import os
import socket

def test_token():
    # Load config
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    with open(config_path, 'r') as f:
        config = json.load(f)
    
    token = config['joplin_api_token']
    base_url = config['web_clipper']['url']
    port = config['web_clipper']['port']
    
    print(f"Testing connection to {base_url}:{port}...")
    
    # First check if the port is open
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex(('127.0.0.1', port))
        sock.close()
        
        if result != 0:
            print(f"❌ Port {port} is not open. Is Joplin running?")
            return
        print("✅ Port is open")
    except Exception as e:
        print(f"❌ Error checking port: {str(e)}")
        return
    
    # Test the token
    url = f"{base_url}:{port}/notes"
    params = {'token': token}
    
    print("\nTesting API token...")
    try:
        response = requests.get(url, params=params, timeout=5)
        if response.status_code == 200:
            print("✅ Token is valid!")
            print(f"Response: {response.json()}")
        else:
            print(f"❌ Token validation failed with status code: {response.status_code}")
            print(f"Response: {response.text}")
    except requests.exceptions.ConnectionError:
        print("❌ Could not connect to Joplin. Is the Web Clipper service enabled?")
    except Exception as e:
        print(f"❌ Error testing token: {str(e)}")

if __name__ == "__main__":
    test_token() 