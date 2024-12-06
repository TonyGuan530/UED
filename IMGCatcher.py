import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
import os
import threading
from PIL import Image
import json
import ctypes

# Function to capture screenshots of divs and save them to a local folder
def capture_div_screenshots(url, div_class, output_folder, num_screenshots, crop_left, crop_right, crop_top, crop_bottom, log_text):
    try:
        # Set up the Edge driver
        options = webdriver.EdgeOptions()
        options.add_argument('--headless=new')  # Run in headless mode
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument('--log-level=3')
        service = Service(EdgeChromiumDriverManager().install(), log_output=os.devnull)
        driver = webdriver.Edge(service=service, options=options)
        
        log_text.insert(tk.END, "Starting EdgeDriver...\n")
        
        # Open the webpage
        log_text.insert(tk.END, f"Opening URL: {url}\n")
        driver.get(url)
        
        # Wait for the page to load
        log_text.insert(tk.END, "Waiting for the page to load...\n")
        time.sleep(5)
        
        # Find all div elements with the specified class
        log_text.insert(tk.END, f"Finding div elements with class: {div_class}\n")
        divs = driver.find_elements(By.CSS_SELECTOR, f'div.{div_class}')
        log_text.insert(tk.END, f"Found {len(divs)} div elements.\n")
        
        # Create the output folder if it doesn't exist
        if not os.path.exists(output_folder):
            log_text.insert(tk.END, f"Creating output folder: {output_folder}\n")
            os.makedirs(output_folder)
        
        # Capture screenshots of the specified number of divs
        for i in range(min(num_screenshots, len(divs))):
            div = divs[i]
            screenshot_path = os.path.join(output_folder, f'screenshot_{i+1}.png')
            log_text.insert(tk.END, f"Capturing screenshot of div {i+1} and saving to {screenshot_path}\n")
            div.screenshot(screenshot_path)
            log_text.insert(tk.END, f"Screenshot saved to {screenshot_path}\n")
            
            # Crop the screenshot if needed
            if crop_left > 0 or crop_right > 0 or crop_top > 0 or crop_bottom > 0:
                img = Image.open(screenshot_path)
                width, height = img.size
                cropped_img = img.crop((crop_left, crop_top, width - crop_right, height - crop_bottom))
                cropped_img.save(screenshot_path)
                log_text.insert(tk.END, f"Cropped screenshot saved to {screenshot_path}\n")
        
        # Close the browser
        log_text.insert(tk.END, "Closing the browser...\n")
        driver.quit()
        log_text.insert(tk.END, "Browser closed.\n")
        
        # Open the output folder
        os.startfile(output_folder)
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")

# Function to start the screenshot capture process in a separate thread
def start_capture():
    try:
        url = url_entry.get()
        num_screenshots = int(num_screenshots_entry.get())
        output_folder = output_folder_entry.get()
        
        crop_left = int(crop_left_entry.get()) if crop_left_entry.get() else 0
        crop_right = int(crop_right_entry.get()) if crop_right_entry.get() else 0
        crop_top = int(crop_top_entry.get()) if crop_top_entry.get() else 0
        crop_bottom = int(crop_bottom_entry.get()) if crop_bottom_entry.get() else 0
        
        settings = {
            "url": url,
            "num_screenshots": num_screenshots,
            "output_folder": output_folder,
            "crop_left": crop_left,
            "crop_right": crop_right,
            "crop_top": crop_top,
            "crop_bottom": crop_bottom
        }
        
        with open('settings.json', 'w') as f:
            json.dump(settings, f)
        
        log_text.delete(1.0, tk.END)
        
        threading.Thread(target=capture_div_screenshots, args=(url, 'main-container', output_folder, num_screenshots, crop_left, crop_right, crop_top, crop_bottom, log_text)).start()
    except ValueError as e:
        messagebox.showerror("Input Error", "Please enter valid integer values for number of screenshots and cropping dimensions.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to browse for output folder
def browse_output_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        output_folder_entry.delete(0, tk.END)
        output_folder_entry.insert(0, folder_selected)

# Function to load settings from JSON file
def load_settings():
    if os.path.exists('settings.json'):
        with open('settings.json', 'r') as f:
            settings = json.load(f)
            
            url_entry.insert(0, settings.get("url", ""))
            num_screenshots_entry.insert(0, settings.get("num_screenshots", ""))
            output_folder_entry.insert(0, settings.get("output_folder", ""))
            crop_left_entry.insert(0, settings.get("crop_left", ""))
            crop_right_entry.insert(0, settings.get("crop_right", ""))
            crop_top_entry.insert(0, settings.get("crop_top", ""))
            crop_bottom_entry.insert(0, settings.get("crop_bottom", ""))

# Hide the console window (Windows only)
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# Create the main window
root = tk.Tk()
root.title("Screenshot Capture Tool")

# URL label and entry
tk.Label(root, text="URL:").grid(row=0, column=0, padx=10, pady=5)
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=0, column=1, padx=10, pady=5)

# Number of screenshots label and entry
tk.Label(root, text="Number of Screenshots:").grid(row=1, column=0, padx=10, pady=5)
num_screenshots_entry = tk.Entry(root, width=10)
num_screenshots_entry.grid(row=1, column=1, padx=10, pady=5)

# Output folder label and entry with browse button
tk.Label(root, text="Output Folder:").grid(row=2, column=0, padx=10, pady=5)
output_folder_entry = tk.Entry(root, width=50)
output_folder_entry.grid(row=2, column=1, padx=10, pady=5)
browse_button = tk.Button(root, text="Browse", command=browse_output_folder)
browse_button.grid(row=2, column=2, padx=10, pady=5)

# Expandable frame for cropping options
crop_frame = tk.LabelFrame(root, text="Cropping Options", padx=10, pady=10)
crop_frame.grid(row=3, columnspan=3, padx=10, pady=5)

# Crop left label and entry
tk.Label(crop_frame, text="Crop Left (pixels):").grid(row=0, column=0, padx=10, pady=5)
crop_left_entry = tk.Entry(crop_frame, width=10)
crop_left_entry.grid(row=0, column=1)

# Crop right label and entry
tk.Label(crop_frame, text="Crop Right (pixels):").grid(row=1, column=0, padx=10, pady=5)
crop_right_entry = tk.Entry(crop_frame, width=10)
crop_right_entry.grid(row=1, column=1)

# Crop top label and entry
tk.Label(crop_frame, text="Crop Top (pixels):").grid(row=2, column=0, padx=10, pady=5)
crop_top_entry = tk.Entry(crop_frame, width=10)
crop_top_entry.grid(row=2, column=1)

# Crop bottom label and entry
tk.Label(crop_frame, text="Crop Bottom (pixels):").grid(row=3, column=0, padx=10, pady=5)
crop_bottom_entry = tk.Entry(crop_frame, width=10)
crop_bottom_entry.grid(row=3, column=1)

# Log text box
log_text = tk.Text(root, height=15, width=80)
log_text.grid(row=4, column=0, columnspan=3, padx=10, pady=5)

# Capture button
capture_button = tk.Button(root, text="Capture Screenshots", command=start_capture)
capture_button.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

# Load settings on startup
load_settings()

# Start the Tkinter event loop
root.mainloop()

# Close the console window after the program finishes
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)