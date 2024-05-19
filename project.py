import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re  # Import the regex module
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import matplotlib.pyplot as plt
from pandas.plotting import table

def setup_driver():
    options = Options()
    options.use_chromium = True
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    service = Service(r'C:\edgedriver\msedgedriver.exe')
    driver = webdriver.Edge(service=service, options=options)
    return driver

driver = setup_driver()

def accept_cookies():
    try:
        wait = WebDriverWait(driver, 10)
        accept_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="dialog-accept-button"]')))
        accept_button.click()
    except Exception:
        print("Failed to accept cookies.")

def clean_price(price_str):
    # Remove any non-numeric characters except decimal point
    cleaned_price = re.sub(r'[^\d.]', '', price_str)
    return float(cleaned_price)

def retrieve_data():
    url = 'https://www.nike.com/ro/w/promotions-skateboarding-shoes-8mfrfz9dklkzy7ok'
    driver.get(url)
    accept_cookies()
    products = [elem.text for elem in driver.find_elements(By.CSS_SELECTOR, 'div.product-card__title')]
    prices = [elem.text for elem in driver.find_elements(By.CSS_SELECTOR, 'div.product-price.is--current-price')]
    prices = [clean_price(price) for price in prices]  # Clean each price using the clean_price function
    return pd.DataFrame({'Product': products, 'Price': prices}), True

def create_graph(data):
    if data is None:
        return
    plt.figure(figsize=(10, 5))
    plt.bar(data['Product'], data['Price'])
    plt.xlabel('Product')
    plt.ylabel('Price')
    plt.title('Product Prices')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

def display_matrix(data):
    if data is None:
        return
    fig, ax = plt.subplots(figsize=(12, 2))
    ax.xaxis.set_visible(False)
    ax.yaxis.set_visible(False)
    ax.set_frame_on(False)
    tab = table(ax, data, loc='upper center', colWidths=[0.17]*len(data.columns))
    tab.auto_set_font_size(False)
    tab.set_fontsize(12)
    tab.scale(1.2, 1.2)
    plt.show()

def save_to_excel(data, filename):
    if data is None:
        return
    filepath = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')], initialfile=filename)
    data.to_excel(filepath, index=False)
    print(f'Saved to {filepath}')

root = tk.Tk()
root.state('zoomed')

header = tk.Label(root, text="PROJECT SCRAPPING", font=("Helvetica", 14), fg="white", bg="#12447d")
header.pack(fill=tk.X, padx=10, pady=10)

filename_var = tk.StringVar()
filename_entry = tk.Entry(root, textvariable=filename_var)
filename_entry.pack(pady=10)

data = None

def update_data():
    global data
    data, success = retrieve_data()
    if success:
        retrieve_button.config(bg='green')
        messagebox.showinfo("Status", "Data loaded successfully!", parent=root)
    else:
        retrieve_button.config(bg='blue')
        messagebox.showinfo("Status", "Failed to load data!", parent=root)

button_style = {'font': ("Helvetica", 14), 'bg': '#4F46E5', 'fg': 'white', 'padx': 10, 'pady': 5}
retrieve_button = tk.Button(root, text='Retrieve Data', command=update_data, **button_style)
retrieve_button.pack(pady=10)

tk.Button(root, text='Create Graph', command=lambda: create_graph(data), **button_style).pack(pady=10)
tk.Button(root, text='Display Matrix', command=lambda: display_matrix(data), **button_style).pack(pady=10)
tk.Button(root, text='Save to Excel', command=lambda: save_to_excel(data, filename_entry.get()), **button_style).pack(pady=10)

root.mainloop()
driver.quit()
