import keyboard
import time
import pygetwindow as gw
from PIL import ImageGrab, Image
from io import BytesIO
import win32clipboard  # pywin32

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Set up Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--use-fake-ui-for-media-stream") # allows mic
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://www.bing.com/search?form=NTPCHB&q=Bing+AI&showconv=1")


# Function to find the "Control microphone" button by its class name
def find_mic_button():
    mic_button = None
    reach_micbutton_when_searchbox_focus_script = """
        return document.querySelector("#b_sydConvCont > cib-serp").shadowRoot.querySelector("#cib-action-bar-main").shadowRoot.querySelector("div > div.main-container > div > div.input-row > div > div > button");
     """
    # reach_micbutton_when_searchbox_outfocus_script = """
    #     return document.querySelector("#b_sydConvCont > cib-serp").shadowRoot.querySelector("#cib-action-bar-main").shadowRoot.querySelector("#cib-speech-icon").shadowRoot.querySelector("button")
    # """
    try:
        find_search_box().click() # must focus the searchbox before clicking mic button
        mic_button = driver.execute_script(reach_micbutton_when_searchbox_focus_script)

    except Exception as e:
        print(e)

    return mic_button


# Function to find the search box by its ID using XPath
def find_search_box():
    try:
        reach_searchbox_script = """
            return document.querySelector("#b_sydConvCont > cib-serp").shadowRoot.querySelector("#cib-action-bar-main").shadowRoot.querySelector("div > div.main-container > div > div.input-row > cib-text-input").shadowRoot.querySelector("#searchbox")
        """
        chat_box_element = driver.execute_script(reach_searchbox_script)
        return chat_box_element
    except Exception as e:
        print(f"Error finding search box: {e}")
        return None


# Function to take a screenshot of the active window and save it to the clipboard
def take_screenshot():
    active_window = gw.getActiveWindow()
    if active_window:
        window_rect = active_window.box
        screenshot = ImageGrab.grab(window_rect)

        # Save the image to the clipboard
        output = BytesIO()
        screenshot.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()


# Function to focus on the search box and paste the screenshot
def paste_screenshot():
    search_box = find_search_box()
    if search_box:
        # Clear existing text in the input field
        search_box.clear()
        # Use send_keys to insert clipboard contents directly
        search_box.send_keys(Keys.CONTROL, 'v')
        # Desc for the img
        search_box.send_keys("Help me.")  # todo variable
        # Wait uploading
        time.sleep(3)
        # Submit
        search_box.send_keys(Keys.ENTER)


# Callback function for key events
def on_key_event(e):
    try:
        if e.event_type == keyboard.KEY_DOWN:
            if e.name.lower() == "r":
                mic_button = find_mic_button()
                if mic_button:
                    time.sleep(0.1)
                    mic_button.click()  # Toggle the "Control microphone" button
            elif e.name.lower() == "s":
                take_screenshot()
                time.sleep(0.5)
                paste_screenshot()
    except Exception as e:
        print(f"Error: {e}")


# Register the callback function for key events
keyboard.hook(on_key_event)

try:
    # Keep the script running
    keyboard.wait("esc")
finally:
    # Clean up resources
    keyboard.unhook_all()
    driver.quit()
