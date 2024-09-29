import os
import subprocess
import requests
import time
import sys
import ctypes
import tempfile
import sounddevice as sd
from pywinauto import Application
from pywinauto.timings import wait_until
from zipfile import ZipFile, BadZipFile
from io import BytesIO

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if not is_admin():
        print("Run the script as administrator")
        sys.exit()

def is_default_mic_stereo():
    try:
        # Get the default input device index
        default_device_index = sd.default.device[0]
        # Get device info
        device_info = sd.query_devices(default_device_index)
        # Check the 'max_input_channels'
        return device_info['max_input_channels'] == 2
    except Exception as e:
        print(f"Error: {e}")
        return None

def get_latest_github_release_url(repo, file_prefix, pre_release=False):
    """
    Gets the URL of the latest release ZIP file that contains a specified prefix in its name from a GitHub repository.
    
    :param repo: The GitHub repository in the format 'owner/repo'.
    :param file_prefix: The prefix to look for in the ZIP file name.
    :param pre_release: Boolean indicating whether to include pre-releases.
    :return: The URL of the latest release ZIP file that contains the specified prefix, or None if not found.
    """
    try:
        # GitHub API URL for the latest releases
        api_url = f"https://api.github.com/repos/{repo}/releases"
        
        # Send a GET request to the API URL
        response = requests.get(api_url)
        response.raise_for_status()
        
        # Parse the JSON response
        releases = response.json()
        
        # Find the latest release (including pre-releases if specified)
        latest_release = None
        for release in releases:
            if pre_release or not release['prerelease']:
                latest_release = release
                break
        
        if not latest_release:
            print("No releases found.")
            return None
        
        # Find the asset that contains the specified prefix
        for asset in latest_release['assets']:
            if file_prefix in asset['name'] and asset['content_type'] == 'application/zip':
                return asset['browser_download_url']
        
        print(f"No ZIP file containing '{file_prefix}' found in the latest release.")
        return None
    
    except requests.RequestException as e:
        print(f"HTTP request failed: {e}")
        return None

def download_and_extract_zip(url, extract_to):
    """
    Downloads a ZIP file from a URL and extracts it to a specified directory.
    
    :param url: The URL of the ZIP file.
    :param extract_to: The directory to extract the contents to.
    """
    try:
        # Send a GET request to the URL
        response = requests.get(url)
        
        # Raise an HTTPError if the HTTP request returned an unsuccessful status code
        response.raise_for_status()
        
        # Create a ZipFile object from the response content
        with ZipFile(BytesIO(response.content)) as zip_file:
            # Extract all the contents to the specified directory
            zip_file.extractall(extract_to)
        
        print(f"Files extracted to {extract_to}")
    
    except requests.RequestException as e:
        print(f"HTTP request failed: {e}")
    except BadZipFile as e:
        print(f"Error with the ZIP file: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Step 1: Download the Equalizer APO installer
def download_equalizer_apo(installer_url, download_path):
    response = requests.get(installer_url)
    with open(download_path, 'wb') as file:
        file.write(response.content)
    print(f"Downloaded Equalizer APO installer to {download_path}")

# Step 2: Install Equalizer APO silently
def install_equalizer_apo(installer_path):
    subprocess.Popen([installer_path, '/S'])
    print("Started Equalizer APO installer")

def get_app_window(title_re):
    # Wait for the app to launch
    app = None
    for _ in range(20):  # Retry for a maximum of 20 seconds
        try:
            app = Application(backend="uia").connect(title_re=title_re)
            break
        except:
            time.sleep(1)

    if app is None:
        print(f"Could not connect to app with regex {title_re}")
        return

    time.sleep(0.1)  # Wait for the app to fully open

    # Interact with the app GUI
    dialog = app.window(title_re=title_re)  # Use a more flexible regex
    return dialog

# Step 3: Run DeviceSelector.exe and interact with the GUI to configure the microphone
def run_equalizer_apo_device_selector(microphone_device):
    dialog = get_app_window(".*Device Selector.*")
    # dialog.print_control_identifiers()  # Print control identifiers for debugging

    # Wait until the TreeView is ready to be interacted with
    tree_view = dialog.child_window(control_type="Tree")
    wait_until(5, 0.5, lambda: tree_view.exists())

    # Get all items from the TreeView
    children = tree_view.children()
    text_items = [i.window_text() for i in children]

    mics = False
    changed = False
    for i in range(len(children)):
        if text_items[i] == "Playback devices":
            children[i].select()
            children[i].click_input(double=True)

        # skip until we get to capture devices
        if not mics:
            mics = bool(text_items[i] == "Capture devices")
            continue

        if microphone_device in text_items[i]:
            j = (i//3)*3
            connector, device, status = text_items[j], text_items[j+1], text_items[j+2]
            microphone_device = connector + " " + device

            microphone_item = children[j]
            
            if not "is already installed" in status:
                microphone_item.select()  # Attempt to select the TreeView to bring items into view

                time.sleep(0.1)
                microphone_item.click_input()

                # Wait a moment to ensure it's selected
                time.sleep(0.1)

                # Use type_keys to send the spacebar press
                microphone_item.type_keys(' ', with_spaces=True)

                changed = True
                print("Microphone checkbox checked.")
            else:
                print("APO is already installed")

            found = True
            break

    if not found:
        print(f"Microphone device '{microphone_device}' not found in the list")

    if changed:
        dialog.OK.click()
    else:
        # Find the Close button, specify index if there are duplicates
        close_button = dialog.child_window(title="Close", control_type="Button", found_index=0)  # 0 for the first occurrence
        close_button.click()

    print(f"Equalizer APO configured for {microphone_device}")
    
    dialog = get_app_window(".*Testing APO.*")
    dialog.OK.click()

    dialog = get_app_window(".*Info.*")
    dialog.OK.click()
    
    print(f"Closed Equalizer APO")

    return microphone_device

def write_specific_string(file_path, specific_string_to_add, default_string=None):
    # Check if the file exists
    try:
        with open(file_path, 'x') as file:  # 'x' mode raises an error if the file exists
            file.write(specific_string_to_add + '\n')
            print(f'File created and string written: \n{specific_string_to_add}')
    except FileExistsError:
        # If the file exists, read the existing content
        with open(file_path, 'r') as file:
            existing_content = file.read()

        # Check if the existing content matches the default string if provided
        if default_string and existing_content == default_string:
            # Write new contents as if it were a new file
            with open(file_path, 'w') as file:  # 'w' mode to overwrite the file
                file.write(specific_string_to_add + '\n')
                print(f'File overwritten with new string: \n{specific_string_to_add}')
        else:
            # Append the specific string if it is not already present
            if specific_string_to_add not in existing_content:
                # Check if the last character is a newline and add one if not
                with open(file_path, 'a') as file:  # 'a' mode to append to the file
                    if not existing_content.endswith('\n'):
                        file.write('\n')
                    file.write(specific_string_to_add + '\n')
                    print(f'String added: \n{specific_string_to_add}')
            else:
                print(f'String already exists in the file: \n{specific_string_to_add}')

# Step 4: Configure Equalizer APO for the microphone
def configure_equalizer_apo(microphone_device):
    config_path = os.path.join(os.getenv('ProgramFiles'), 'EqualizerAPO', 'config', 'config.txt')

    # get stereo or mono for rnnoise
    audio_channel = "stereo" if is_default_mic_stereo() else "mono"
    print(f"Microphone detected as {audio_channel}")
    
    lines = [
        f'Device: {microphone_device}',
        f'VSTPlugin: Library LoudMax64.dll Thresh 0 Output 0.665909 "Fader Link" 0 "ISP Detection" 0 "Large GUI" 0',
        f'VSTPlugin: Library win-rnnoise\\vst\\rnnoise_{audio_channel}.dll ChunkData "VkMyIdEAAAA8P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJVVEYtOCI/PiA8Uk5Ob2lzZT48UEFSQU0gaWQ9InZhZF9ncmFjZV9wZXJpb2QiIHZhbHVlPSIyMC4wIi8+PFBBUkFNIGlkPSJ2YWRfcmV0cm9hY3RpdmVfZ3JhY2VfcGVyaW9kIiB2YWx1ZT0iMC4wIi8+PFBBUkFNIGlkPSJ2YWRfdGhyZXNob2xkIiB2YWx1ZT0iMC42NDk5OTk5NzYxNTgxNDIxIi8+PC9STk5vaXNlPgA="'
    ]

    # default config.txt file for EqualizerAPO (for overwriting)
    default = 'Preamp: -6 dB\nInclude: example.txt\nGraphicEQ: 25 0; 40 0; 63 0; 100 0; 160 0; 250 0; 400 0; 630 0; 1000 0; 1600 0; 2500 0; 4000 0; 6300 0; 10000 0; 16000 0'

    write_specific_string(config_path, "\n".join(lines), default)
    print(f"Equalizer APO configured for {microphone_device}")

def main():
    print("\nDO NOT TYPE OR MOVE MOUSE WHILE PROGRAM IS RUNNING\n")

    installer_url = "https://sourceforge.net/projects/equalizerapo/files/latest/download"
    download_file = "EqualizerAPO-setup.exe"
    loudmax_url = "https://www.dropbox.com/scl/fi/yovjswlx94m7u6qink5sk/LoudMax_v1_45_WIN_VST2.zip?rlkey=tjjc50g4h120n8jxf1iud6qyc&dl=1"
    rnn_noise_url = get_latest_github_release_url("werman/noise-suppression-for-voice", "win-")
    vst_directory = "C:/Program Files/EqualizerAPO/VSTPlugins"
    microphone_device = "Default"

    with tempfile.TemporaryDirectory() as temp_dir:
        # Create the full path for the new file
        temp_file_path = os.path.join(temp_dir, download_file)

        download_equalizer_apo(installer_url, temp_file_path)
        install_equalizer_apo(temp_file_path)
        microphone_device = run_equalizer_apo_device_selector(microphone_device)
        
        download_and_extract_zip(loudmax_url, vst_directory)
        download_and_extract_zip(rnn_noise_url, vst_directory)
        configure_equalizer_apo(microphone_device)
    
    print("\nFINISHED. Closing in 3 seconds...\n")
    time.sleep(3)

if __name__ == "__main__":
    run_as_admin()  # Ensure the script runs with administrator privileges
    main()
