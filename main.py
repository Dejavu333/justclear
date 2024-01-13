import keyboard
import time
import pygetwindow as gw
from PIL import ImageGrab
from io import BytesIO
import win32clipboard  # pywin32

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Set up Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--use-fake-ui-for-media-stream")  # allows mic
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://www.bing.com/search?form=NTPCHB&q=Bing+AI&showconv=1")


def find_mic_button():
    try:
        reach_micbutton_script = """
            return document.querySelector("#b_sydConvCont > cib-serp").shadowRoot.querySelector("#cib-action-bar-main").shadowRoot.querySelector("div > div.main-container > div > div.input-row > div > div > button")
        """
        mic_button = driver.execute_script(reach_micbutton_script)
        return mic_button
    except Exception as e:
        print(f"Error finding mic button: {e}")
    return None


def focus_searchbox():
    sb = find_search_box()
    sb.click()


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


# take a screenshot of the active window and save it to the clipboard
def print_screen():
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


# focus on the search box and paste the screenshot
def paste_screenshot():
    search_box = find_search_box()
    if search_box:
        # Use send_keys to insert clipboard contents directly
        search_box.send_keys(Keys.CONTROL, 'v')
        # Desc for the img
        search_box.send_keys("Help me.")  # todo variable
        # Wait uploading
        time.sleep(2)


def submit_input():
    find_search_box().send_keys(Keys.ENTER)


def clear_input():
    find_search_box().clear()
    reach_dismiss_button_script = """
        return document.querySelector("#b_sydConvCont > cib-serp").shadowRoot.querySelector("#cib-action-bar-main").shadowRoot.querySelector("div > div.main-container > div > cib-attachment-list").shadowRoot.querySelector("cib-file-item").shadowRoot.querySelector("button");
    """
    btn = driver.execute_script(reach_dismiss_button_script)
    if btn:
        btn.click()


# Callback function for key events
def on_key_event(e):
    try:
        if e.event_type == keyboard.KEY_DOWN:
            # record
            # --------------------------------
            if e.name == "3" and isNumpad(e):
                focus_searchbox() # must do before click on mic button
                mic_button = find_mic_button()
                if mic_button:
                    time.sleep(0.1)
                    mic_button.click()
            # focus / stop record
            # --------------------------------
            elif e.name == "2" and isNumpad(e):
                focus_searchbox()  # focusing the searchbox stops recording too
            # clear
            # --------------------------------
            elif e.name == "1" and isNumpad(e):
                clear_input()
            # submit
            # --------------------------------
            elif e.name == 'enter' and isNumpad(e):
                submit_input()
            # prtsc
            # --------------------------------
            elif keyboard.is_pressed('del'):
                print_screen()
                time.sleep(0.5)
                paste_screenshot()
            # Reload the page
            # --------------------------------
            elif e.name == "7" and isNumpad(e):
                driver.refresh()
    except Exception as e:
        print(f"Error: {e}")


def isNumpad(e):
    return e.event_type == keyboard.KEY_DOWN and e.is_keypad


# Register the callback function for key events
keyboard.hook(on_key_event)

try:
    # Keep the script running
    keyboard.wait("esc")
finally:
    # Clean up resources
    keyboard.unhook_all()
    driver.quit()
