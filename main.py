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


########################################################################################################################
# functions
########################################################################################################################
def record():
    reach_mic_button_script1 = """
        return document.querySelector("#b_sydConvCont > cib-serp").shadowRoot.querySelector("#cib-action-bar-main").shadowRoot.querySelector("div > div.main-container > div > div.input-row > div > div > button")
    """
    reach_mic_button_script2 = """
        return document.querySelector("#b_sydConvCont > cib-serp").shadowRoot.querySelector("#cib-action-bar-main").shadowRoot.querySelector("#cib-speech-icon").shadowRoot.querySelector("button")
    """
    try:
        driver.execute_script(reach_mic_button_script1).click()
    except:
        print(f"Error clicking mic button using script1.")
    try:
        driver.execute_script(reach_mic_button_script2).click()
    except:
        print(f"Error clicking mic button using script1.")

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


def print_screen_to_clipboard():
    active_window = gw.getActiveWindow()
    if active_window:
        window_rect = active_window.box
        screenshot = ImageGrab.grab(window_rect)

        # saves the image to the clipboard
        output = BytesIO()
        screenshot.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()


def paste_screenshot_to_searchbox():
    search_box = find_search_box()
    if search_box:
        # inserts clipboard contents directly
        search_box.send_keys(Keys.CONTROL, 'v')
        # desc for the img
        search_box.send_keys("Help me.")  # todo variable


def submit_input():
    find_search_box().send_keys(Keys.ENTER)


def clear_input():
    find_search_box().clear()
    reach_dismiss_button_script = """
        return document.querySelector("#b_sydConvCont > cib-serp").shadowRoot.querySelector("#cib-action-bar-main").shadowRoot.querySelector("div > div.main-container > div > cib-attachment-list").shadowRoot.querySelector("cib-file-item").shadowRoot.querySelector("button");
    """
    try:
        driver.execute_script(reach_dismiss_button_script).click()
    except:
        print("There was no image to delete.")


def on_key_event(e):
    try:
        if e.event_type == keyboard.KEY_DOWN:
            # record
            # --------------------------------
            if e.name == "3" and is_numpad(e):
                record()
            # focus / stop record
            # --------------------------------
            elif e.name == "2" and is_numpad(e):
                focus_searchbox()  # focusing the searchbox stops recording too
            # clear
            # --------------------------------
            elif e.name == "1" and is_numpad(e):
                clear_input()
            # submit
            # --------------------------------
            elif e.name == 'enter' and is_numpad(e):
                submit_input()
            # prtsc
            # --------------------------------
            elif keyboard.is_pressed('del'):
                print_screen_to_clipboard()
                time.sleep(0.5)
                paste_screenshot_to_searchbox()
            # page reload
            # --------------------------------
            elif e.name == "7" and is_numpad(e):
                driver.refresh()
    except Exception as e:
        print(f"Error: {e}")


def is_numpad(e):
    return e.event_type == keyboard.KEY_DOWN and e.is_keypad


########################################################################################################################
# main
########################################################################################################################
# sets up Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--use-fake-ui-for-media-stream")  # allows mic
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://www.bing.com/search?form=NTPCHB&q=Bing+AI&showconv=1")

# registers the callback function for key events
keyboard.hook(on_key_event)

try:
    # keeps the script running
    keyboard.wait("esc")
finally:
    # cleans up resources
    keyboard.unhook_all()
    driver.quit()
