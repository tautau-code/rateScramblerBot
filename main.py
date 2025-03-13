import pyautogui
import time
from PIL import Image
import os
import ctypes
from datetime import datetime
import glob
import keyboard
import win32api
import win32con

def block_input():
    """Blocks mouse input"""
    return win32api.BlockInput(True)

def unblock_input():
    """Unblocks mouse input"""
    return win32api.BlockInput(False)

def check_keyboard_layout():
    """Checks current keyboard layout and switches to English if needed"""
    # Get active window handle
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    # Get keyboard layout ID
    thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, 0)
    layout_id = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    # Check if layout is English (0x409 is English layout code)
    if layout_id & 0xFFFF != 0x409:
        # Simulate Alt+Shift to switch layout
        pyautogui.hotkey('alt', 'shift')
        time.sleep(0.25)

def click_at_coordinates(x, y):
    """Clicks mouse at given coordinates"""
    pyautogui.moveTo(x, y)
    time.sleep(0.15)  # Reduced delay before click
    pyautogui.click(x, y)

def type_text(text):
    """Types given text"""
    pyautogui.typewrite(text)

def type_numbers(numbers):
    """Types sequence of numbers"""
    pyautogui.typewrite(str(numbers))

def take_screenshot(filename):
    """Takes screenshot of entire screen and saves to file"""
    screenshot = pyautogui.screenshot()
    screenshot.save(filename)

def take_region_screenshot(x, y, width, height, filename):
    """Takes screenshot of specified screen region and saves to file"""
    screenshot = pyautogui.screenshot(region=(x, y, width, height))
    screenshot.save(filename)

def format_time(seconds):
    """Formats time in days, hours and minutes"""
    days = int(seconds // (24 * 3600))
    seconds %= (24 * 3600)
    hours = int(seconds // 3600)
    seconds %= 3600
    minutes = int(seconds // 60)
    time_str = ""
    if days > 0:
        time_str += f"{days}d "
    if hours > 0 or days > 0:
        time_str += f"{hours}h "
    time_str += f"{minutes}min"
    return time_str

def get_last_screenshot_info():
    """Finds last screenshot folder and last screenshot"""
    # Find all screenshot folders
    screenshot_folders = glob.glob("screenshots_*")
    if not screenshot_folders:
        return None, None
    
    # Get newest folder
    latest_folder = max(screenshot_folders)
    
    # Find all screenshots in folder
    screenshots = glob.glob(os.path.join(latest_folder, "screenshot_*.png"))
    if not screenshots:
        return latest_folder, None
    
    # Get latest screenshot
    latest_screenshot = max(screenshots)
    
    # Extract item names from filename
    filename = os.path.basename(latest_screenshot)
    items = filename.replace("screenshot_", "").replace(".png", "").split("_")
    
    return latest_folder, items

if __name__ == "__main__":
    # Check for interrupted session
    last_folder, last_items = get_last_screenshot_info()
    
    if last_folder and last_items:
        screenshots_dir = last_folder
        print(f"Found interrupted session in folder {last_folder}")
        print(f"Continuing with pair after {last_items[0]} -> {last_items[1]}")
    else:
        # Create new folder for current run
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        screenshots_dir = f"screenshots_{current_time}"
        os.makedirs(screenshots_dir, exist_ok=True)

    # Set coordinates manually
    coordinates = {
        'i_want': (903, 325),
        'search': (1253, 1253), 
        'select': (1414, 256),
        'i_have': (1555, 322)
    }

    # Check keyboard layout
    check_keyboard_layout()

    # 2.5 second pause before start for preparation
    time.sleep(2.5)

    # Read items list from file
    with open('items.txt', 'r', encoding='utf-8') as f:
        items = [line.strip() for line in f.readlines()]

    # Generate item pairs (without identical pairs)
    pairs = [(items[i], items[j]) for i in range(len(items)) for j in range(len(items)) if i != j]

    # If interrupted session exists, skip already processed pairs
    if last_items:
        for i, pair in enumerate(pairs):
            if pair[0] == last_items[0] and pair[1] == last_items[1]:
                pairs = pairs[i+1:]
                break

    # Calculate total pairs and start time
    total_pairs = len(pairs)
    start_time = time.time()

    # Save remaining pairs to file
    pairs_file = os.path.join(screenshots_dir, 'item_pairs.txt')
    if not last_items:  # Only for new session
        with open(pairs_file, 'w', encoding='utf-8') as f:
            for pair in pairs:
                f.write(f"{pair[0]},{pair[1]}\n")

    # Confirmation before starting process
    confirm = input("Are you sure you want to start the script? (yes/no): ")
    if confirm.lower() != 'yes':
        print("Operation cancelled.")
        exit()
    # Check keyboard layout
    check_keyboard_layout()

    print("\nPress Backspace to stop the script")
    # 2.5 second pause before start for preparation
    time.sleep(2.5)

    # Block mouse input
    block_input()

    try:
        last_i_want = None
        last_i_have = None

        # Read pairs from file and perform actions for each
        for idx, pair in enumerate(pairs, 1):
            # Check for Backspace key press
            if keyboard.is_pressed('backspace'):
                print("\nScript stopped by user")
                break
                
            # Check layout before entering each pair
            check_keyboard_layout()
            
            # Calculate progress and remaining time
            progress = idx / total_pairs
            elapsed_time = time.time() - start_time
            estimated_total_time = elapsed_time / progress if progress > 0 else 0
            remaining_time = estimated_total_time - elapsed_time

            # Display progress and time
            print(f"\rProgress: [{('='*int(50*progress)):50s}] {progress*100:.1f}%", end='')
            print(f" | Elapsed: {format_time(elapsed_time)} | Remaining: {format_time(remaining_time)}", end='')

            i_have_item, i_want_item = pair

            # Check and update only changed items
            if i_want_item != last_i_want:
                click_at_coordinates(*coordinates['i_want'])
                time.sleep(0.5)
                click_at_coordinates(*coordinates['search'])
                type_text(i_want_item)
                time.sleep(0.5)
                click_at_coordinates(*coordinates['select'])
                time.sleep(0.5)
                last_i_want = i_want_item

            if i_have_item != last_i_have:
                click_at_coordinates(*coordinates['i_have'])
                time.sleep(0.5)
                click_at_coordinates(*coordinates['search'])
                type_text(i_have_item)
                time.sleep(0.5)
                click_at_coordinates(*coordinates['select'])
                time.sleep(0.5)
                last_i_have = i_have_item
            
            # Reduced pause for exchange rate loading
            time.sleep(0.5)

            # Take screenshot for each pair
            screenshot_name = os.path.join(screenshots_dir, f"screenshot_{i_have_item}_{i_want_item}.png")
            take_screenshot(screenshot_name)

        # Calculate execution time
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"\nTotal execution time: {elapsed_time/60:.2f} minutes.")

        print("Current working directory:", os.getcwd())

    finally:
        # Always unblock mouse input
        unblock_input()
