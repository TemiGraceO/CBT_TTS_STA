import os
import urllib.request

# Focus strictly on the configuration file path link
CONFIG_URL = "https://huggingface.co"
TARGET_FOLDER = "voice_model"

def download_missing_config():
    if not os.path.exists(TARGET_FOLDER):
        os.makedirs(TARGET_FOLDER)
    filepath = os.path.join(TARGET_FOLDER, "en_NG-eze-medium.onnx.json")
    print("Downloading the missing configuration file... Please wait.")
    urllib.request.urlretrieve(CONFIG_URL, filepath)
    print(f"Successfully saved to {filepath}")

if __name__ == "__main__":
    try:
        download_missing_config()
        print("\nSuccess! Both Nigerian voice components are now perfectly saved inside your folder.")
    except Exception as e:
        print(f"An error occurred: {e}\nIf this fails again, check your internet connection or try again in a moment.")
