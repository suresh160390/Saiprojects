import subprocess

def list_wifi_profiles():
    try:
        result = subprocess.check_output(['netsh', 'wlan', 'show', 'profiles']).decode('utf-8')
        print("Wi-Fi profiles:")
        print(result)
    except subprocess.CalledProcessError as e:
        print("Error:", str(e))

if __name__ == "__main__":
    list_wifi_profiles()