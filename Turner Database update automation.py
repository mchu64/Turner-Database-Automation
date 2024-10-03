import fitz
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import threading


driver_path = None
pause_flag = threading.Event()
stop_flag = threading.Event()   
links_to_process = []  
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

def process_links():
    """Process the links using Selenium."""
    global authentication_confirmed, stop_flag
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

        # Navigate to the login page and log in

        # CENSORED LOGIN INFORMATION

        # Open the GUI after logging in
        control_window()

        updates = []

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
            location_row = retry_find_element(driver, By.XPATH, "//tr[th[text()='Location:']]")
            if location_row:
                    location_cell = location_row.find_elements(By.TAG_NAME, "td")[0].text.strip()

                    location = ""
                    if location_cell:
                        if location_cell.startswith("Level"):
                            location = location_cell.split("Level")[1].strip()
                        else:
                            location = location_cell

                    if "P" in location:
                        print(f"Skipping page {link} as location contains 'P': '{location}'")
                        continue

                    th_element = retry_find_element(driver, By.XPATH, "//th[@class='v-top' and contains(text(), 'Linked Drawings:')]")
                    if th_element:
                        td_element = th_element.find_element(By.XPATH, "following-sibling::td")
                        a_tag = td_element.find_element(By.TAG_NAME, "a")
                        linked_drawings_text = a_tag.text.strip()

                        if not linked_drawings_text:
                                print(f"No linked drawings found on page {link}. Selecting 'No linked drawing' option...")

                                # Click the Edit button
                                edit_button = driver.find_element(By.XPATH, "//form[@class='button_to']//input[@type='submit']")
                                if edit_button:
                                    edit_button.click()
                                    time.sleep(3)  # Wait for the edit page to load

                                    # Select 'No linked drawing' option
                                    dropdown = retry_find_element(driver, By.XPATH, "//select[@name='punch_item[punch_item_type_id]']")
                                    if dropdown:
                                        options = dropdown.find_elements(By.TAG_NAME, 'option')
                                        for option in options:
                                            if option.get_attribute('value') == '1743793':
                                                option.click()
                                                print("Selected 'No linked drawing' option.")
                                                break

                                        save_button = retry_find_element(driver, By.XPATH, "//button[@type='submit' and contains(@class, 'punch-form-submit') and contains(text(), 'Save')]")
                                        if save_button:
                                            save_button.click()
                                            print("Save button clicked.")

                                updates.append({
                                    "URL": link,
                                    "Original Location": location,
                                    "Linked Drawing": linked_drawings_text,
                                    "Updated Location": "No linked drawing"
                                })
                        else:
                                linked_drawings = extract_number_from_linked_drawings(linked_drawings_text)

                                if linked_drawings.startswith('0'):
                                    linked_drawings = linked_drawings[1:]

                                print(f"Location: {location}, Linked Drawings: {linked_drawings}")

                                if location != linked_drawings:
                                    driver.get(link + '/edit')
                                    select = "Level " + linked_drawings

                                    location_button = retry_find_element(driver, By.XPATH, '//*[@id="punch-form"]/section/table[1]/tbody/tr[6]/td[1]/div/div/div')
                                    if location_button:
                                        location_button.click()

                                        search_input = retry_find_element(driver, By.XPATH, "//input[@data-qa='core-typeahead-input']")
                                        if search_input:
                                            search_input.send_keys(select)
                                            time.sleep(2)

                                            dropdown_items = driver.find_elements(By.XPATH, "//div[@data-internal='menuimperative-options']//div[@class='sc-fXoxut caMcxa sc-hCMElv llLZDJ']")

                                            for item in dropdown_items:
                                                item_text = item.text.strip()
                                                print(f"Item: {item_text}")

                                            found = False
                                            for item in dropdown_items:
                                                item_text = item.text.strip()
                                                if item_text == select:
                                                    item.click()
                                                    print(f"Selected item: {item_text}")
                                                    found = True
                                                    break

                                            if not found:
                                                print(f"Could not find the item for '{select}'.")

                                            save_button = retry_find_element(driver, By.XPATH, "//button[@type='submit' and contains(@class, 'punch-form-submit') and contains(text(), 'Save')]")
                                            if save_button:
                                                save_button.click()
                                                print("Save button clicked.")

                                            updates.append({
                                                "URL": link,
                                                "Original Location": location,
                                                "Linked Drawing": linked_drawings,
                                                "Updated Location": select
                                            })
        driver.quit()
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
    """Confirm the authentication process."""
    global authentication_confirmed
    authentication_confirmed = True
    print("Authentication confirmed.")

def setup_gui():
    """Set up the GUI for the application."""
    global root
    root = tk.Tk()
    root.title("PDF Link Extractor")

    # Create GUI elements
    open_button = tk.Button(root, text="Open PDF", command=open_file_dialog)
    open_button.pack(pady=10)

    driver_button = tk.Button(root, text="Select Chromedriver", command=browse_driver_file)
    driver_button.pack(pady=10)

    confirm_button = tk.Button(root, text="Confirm Authentication", command=confirm_authentication)
    confirm_button.pack(pady=10)

    process_button = tk.Button(root, text="Process Links", command=process_links)
    process_button.pack(pady=10)

    exit_button = tk.Button(root, text="Exit", command=root.quit)
    exit_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    setup_gui()