import fitz
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import threading
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation

# Global variables
driver_path = None
pause_flag = threading.Event()
stop_flag = threading.Event()
links_to_process = []
processed_links_count = 0  # For matplotlib progress tracking
authentication_confirmed = False
control_win = None

def extract_hyperlinks(pdf_path):
    """Extract hyperlinks from a PDF file."""
    try:
        pdf_document = fitz.open(pdf_path)
        links = []
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            link_annotations = page.get_links()
            for link in link_annotations:
                if 'uri' in link:
                    uri = link['uri']
                    if 'drawing' not in uri.lower():
                        links.append(uri)
        pdf_document.close()
        return links
    except Exception as e:
        print(f"Error extracting hyperlinks: {e}")
        return []

def extract_number_from_linked_drawings(text):
    """Extract number from linked drawings text."""
    if "P" in text:
        return text.split("P", 1)[1]
    else:
        return text[-2:]

def retry_find_element(driver, by, value, retries=3, delay=2):
    """Retry finding an element with a timeout."""
    for _ in range(retries):
        try:
            element = driver.find_element(by, value)
            if element:
                return element
        except Exception as e:
            print(f"Retry finding element failed: {e}")
        time.sleep(delay)
    raise Exception(f"Element with {by}='{value}' not found after {retries} retries")

def control_window():
    """Open a control window to pause, resume, and stop the script."""
    global control_win

    if control_win and control_win.winfo_exists():
        control_win.lift()
        return

    control_win = tk.Toplevel(root)
    control_win.title("Control Window")
    control_win.geometry("300x150")

    pause_button = tk.Button(control_win, text="Pause", command=pause_script)
    pause_button.pack(pady=5)

    resume_button = tk.Button(control_win, text="Resume", command=resume_script)
    resume_button.pack(pady=5)

    stop_button = tk.Button(control_win, text="Stop", command=stop_script)
    stop_button.pack(pady=5)

def pause_script():
    """Pause the script execution."""
    global pause_flag
    pause_flag.set()
    print("Script paused.")

def resume_script():
    """Resume the script execution."""
    global pause_flag
    pause_flag.clear()
    print("Script resumed.")

def stop_script():
    """Stop the script execution."""
    global stop_flag
    stop_flag.set()
    print("Script stopped.")
    if control_win:
        control_win.destroy()

def toggle_pause():
    """Handle pausing and resuming based on the pause flag."""
    while not stop_flag.is_set():
        if pause_flag.is_set():
            time.sleep(1)

def update_progress_plot(processed, total):
    """Update the progress bar plot."""
    plt.cla()
    plt.barh(['Progress'], [processed], color='green')
    plt.barh(['Progress'], [total - processed], left=[processed], color='gray')
    plt.xlim(0, total)
    plt.xlabel(f'Processed: {processed}/{total}')
    plt.pause(0.1)  # Short pause to allow plot update

def process_links():
    """Process the links using Selenium."""
    global authentication_confirmed, stop_flag, processed_links_count
    if not links_to_process:
        messagebox.showinfo("Info", "No links to process.")
        return

    if driver_path is None:
        messagebox.showerror("Error", "No chromedriver file selected.")
        return

    if not authentication_confirmed:
        messagebox.showerror("Error", "Authentication has not been confirmed.")
        return

    try:
        chrome_options = Options()
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--allow-insecure-localhost')

        service = Service(executable_path=driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)

        pause_thread = threading.Thread(target=toggle_pause, daemon=True)
        pause_thread.start()

        # CENSORED LOGIN INFORMATION

        # Open the GUI after logging in
        control_window()

        total_links = len(links_to_process)
        updates = []

        # Initialize progress bar plot
        plt.ion()
        plt.figure(figsize=(6, 3))

        for link in links_to_process:
            # Check if script is stopped
            if stop_flag.is_set():
                print("Script stopped.")
                break

            # Check if script is paused
            while pause_flag.is_set():
                time.sleep(1)

            # Navigate to the link
            driver.get(link)
            time.sleep(3)

            # Process page as before, skipping drawing checks, etc.

            processed_links_count += 1  # Increment after processing each link

            # Update progress plot
            update_progress_plot(processed_links_count, total_links)

        driver.quit()
        plt.ioff()
    except Exception as e:
        print(f"Error in processing links: {e}")

def open_file_dialog():
    """Open a file dialog to select the PDF file."""
    global driver_path
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        links = extract_hyperlinks(file_path)
        if links:
            global links_to_process
            links_to_process = links
            print(f"Extracted {len(links)} links from PDF.")
        else:
            print("No valid links found in the PDF.")

def browse_driver_file():
    """Open a file dialog to select the chromedriver file."""
    global driver_path
    driver_path = filedialog.askopenfilename(filetypes=[("Executable files", "*.exe")])
    if driver_path:
        print(f"Chromedriver selected: {driver_path}")

def confirm_authentication():
    """Confirm the authentication step."""
    global authentication_confirmed
    authentication_confirmed = True
    messagebox.showinfo("Success", "Authentication confirmed.")
    print("Authentication confirmed.")

# GUI setup
root = tk.Tk()
root.title("PDF Hyperlink Processor")
root.geometry("300x200")

pdf_button = tk.Button(root, text="Select PDF", command=open_file_dialog)
pdf_button.pack(pady=5)

driver_button = tk.Button(root, text="Select Chromedriver", command=browse_driver_file)
driver_button.pack(pady=5)

confirm_button = tk.Button(root, text="Confirm Authentication", command=confirm_authentication)
confirm_button.pack(pady=5)

start_button = tk.Button(root, text="Start Processing", command=process_links)
start_button.pack(pady=5)

root.mainloop()
