import os
import re
import shutil
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException
from bs4 import BeautifulSoup as Bs
import xlrd
import json
import requests
from urllib.parse import urljoin
import json

# A function to load the configuration from the config.json file
def load_config():
    try:
        with open("config.json", "r") as config_file:
            config = json.load(config_file)
            return config
    except FileNotFoundError:
        print("Файл config.json не знайдено. Будь ласка, створіть його.")
        exit(1)

# Get the login and password from the configuration
config = load_config()
login = config["login"]
password = config["password"]

# Get the current directory
current_directory = os.getcwd()

# File path for the Catalog_XXXX.xls in the current directory
file_path = os.path.join(current_directory, 'Catalog_XXXX.xls')

# Open the workbook
workbook = xlrd.open_workbook(file_path)

# Get a list of names of all letters
sheet_names = workbook.sheet_names()

def get_glass_coding(file_path):

    # Select a sheet
    sheet = workbook.sheet_by_name(sheet_names[1])

    type_glass_coding = {}

    for row_glass_type in range(9, 14):
        glass_type = sheet.row_values(row_glass_type)
        type_glass_cod = glass_type[17]
        type_glass_cod_decoding = glass_type[18]
        type_glass_coding[type_glass_cod] = type_glass_cod_decoding

    for row_glass_color in range(16, 39):
        glass_color = sheet.row_values(row_glass_color)
        color_glass_cod = glass_color[17]
        color_glass_cod_decoding = glass_color[18]
        type_glass_coding[color_glass_cod] = color_glass_cod_decoding

    for row_glass_windshields in range(43, 67):
        glass_windshields = sheet.row_values(row_glass_windshields)
        windshields_glass_cod = glass_windshields[17]
        windshields_glass_cod_decoding = glass_windshields[18]
        type_glass_coding[windshields_glass_cod] = windshields_glass_cod_decoding

    for row_glass_back in range(114, 131):
        glass_back = sheet.row_values(row_glass_back)
        back_glass_cod = glass_back[17]
        back_glass_cod_decoding = glass_back[18]
        type_glass_coding[back_glass_cod] = back_glass_cod_decoding

    for row_body_color_glass in range(201, 217):
        body_color_glass = sheet.row_values(row_body_color_glass)
        body_color_glass_cod = body_color_glass[17]
        body_color_glass_cod_decoding = body_color_glass[18]
        type_glass_coding[body_color_glass_cod] = body_color_glass_cod_decoding

    for row_glass_location in range(244, 258):
        glass_location = sheet.row_values(row_glass_location)
        glass_location_cod = glass_location[17]
        glass_location_cod_decoding = glass_location[18]
        type_glass_coding[glass_location_cod] = glass_location_cod_decoding

    for row_accessories in range(310, 331):
        accessories = sheet.row_values(row_accessories)
        accessories_cod = accessories[17]
        accessories_cod_decoding = accessories[18]
        type_glass_coding[accessories_cod] = accessories_cod_decoding

    return type_glass_coding

# get_glass_coding(file_path)

def custom_replace(input_str, glass_encoding):
    # Add spaces between words that are next to each other without a space
    for key in glass_encoding.keys():
        input_str = re.sub(r'(?<=\S)({})(?=\S)'.format(re.escape(key)), r' \1 ', input_str)

    # Decipher the keys
    for key, value in glass_encoding.items():
        # Create a regular expression pattern to match the key as a whole word
        pattern = r'\b{}\b'.format(re.escape(key))
        # Use re.sub to perform the replacement
        input_str = re.sub(pattern, value, input_str)

    # Remove any extra spaces
    input_str = re.sub(r'\s+', ' ', input_str).strip()

    return input_str

def get_glass_producer(file_path):

    # Select a sheet
    sheet = workbook.sheet_by_name(sheet_names[2])

    glass_producers = []

    for row_index in range(7, sheet.nrows):
        row_values = sheet.row_values(row_index)
        glass_producer = row_values[2]

        glass_producers.append(glass_producer)

    return glass_producers


def format_car_models(input_file, output_file):
    # Open the input file for reading
    with open(input_file, 'r', encoding='utf-8') as f:
        car_models_data = [json.loads(line.strip()) for line in f]

    if not os.path.exists(output_file):

        # Write the data in a new file in the desired format
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(car_models_data, f, indent=2, ensure_ascii=False)

        print(f"Файл {input_file} був перетворений і збережений у файл {output_file} у потрібному форматі.")

    else:
        print(f"Файл {output_file} вже існує.")

def process_XXXX_data(input_file, output_file):
    # Open the input file for reading
    with open(input_file, 'r', encoding='utf-8') as f:
        # Read the entire content of the file
        XXXX_data_str = f.read()

    # Separate JSON objects with a comma and a newline
    XXXX_data_str = XXXX_data_str.replace(']\n[', '')
    XXXX_data_str = XXXX_data_str.replace('\n][', ', ')
    XXXX_data_str = XXXX_data_str.replace('}\n', '}, ')
    XXXX_data_str = XXXX_data_str.replace(', ]', '\n]')

    # save the updated content in a new file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(XXXX_data_str)

    print(f"Файл {input_file} був оброблений і збережений у файлі {output_file} у потрібному форматі.")

    # Download the converted file
    with open(output_file, 'r', encoding='utf-8') as f:
        XXXX_data = json.load(f)

        predefined_producers = get_glass_producer(file_path)
        glass_encoding = get_glass_coding(file_path)

        # Iterate through each JSON object and add the "glass producer" field based on the "name" field
        for obj in XXXX_data:
            if "name" in obj:
                glass_producer = obj["name"]

                # Check if the glass producer is in the predefined list
                matching_producers = [producer for producer in predefined_producers if producer in glass_producer]

                # Set the "glass producer" field based on the match
                if matching_producers:
                    obj["manufacturer"] = matching_producers[0]
                else:
                    obj["manufacturer"] = ''

                if "name" in obj:
                    name_parts = obj["name"].split(', ')
                    obj["title"] = ', '.join(name_parts[1:4])
                    obj["full text"] = ', '.join(name_parts)[:].split("=")[0]
                    obj["type_glass"] = ', '.join(name_parts[1:2])
                    obj["desc"] = ', '.join(name_parts[-3:-1])

                    # Apply custom replacement to "title" and "full text"
                    obj["title"] = custom_replace(obj["title"], glass_encoding) + ' ' + ' '.join(name_parts[:1])
                    obj["full text"] = custom_replace(obj["full text"], glass_encoding)
                    obj["type_glass"] = custom_replace(obj["type_glass"], glass_encoding)
                    obj["note"] = ""
                    obj["desc"] = custom_replace(obj["desc"], glass_encoding)

    # Save the updated content in the same file
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(XXXX_data, f, ensure_ascii=False, indent=2)

    print(f"Оброблені дані збережені у файлі {output_file}.")


def combine_data(XXXX_file, car_models_file, output_file):
    import json

    # Open the files for reading
    with open(car_models_file, 'r', encoding='utf-8') as f1, open(XXXX_file, 'r', encoding='utf-8') as f2:
        # Read data from files
        car_models_data = json.load(f1)
        XXXX_data = json.load(f2)

    # Create a new list for the combined data
    combined_data = []

    # Counter for generating unique IDs
    id_counter = 1

    # Go through each object from the "Car_models_formatted.json" file
    for car_model in car_models_data:
        # Отримаємо код моделі з об'єкту
        model_code = car_model.get("Model code", "")

        # Find all objects from the file "XXXX_formatted.json" with the same code
        matching_XXXX_models = [XXXX_model for XXXX_model in XXXX_data if XXXX_model.get("code") == model_code]

        # Combine the data of the two files if there is a code match
        for XXXX_model in matching_XXXX_models:
            # Check if the "title" key exists in XXXX_model
            if "title" in XXXX_model:
                combined_model = {
                    "id": id_counter,
                    "title": XXXX_model["title"],
                    "full_text": XXXX_model.get("full text", ""),  # Use get() to provide a default value
                    "price": int(XXXX_model.get("price", 0)),
                    "status": "in-stock" if XXXX_model.get("status") == "In stock" else "Out of stock",
                    "quantity": "",  # You can set the quantity as needed
                    "euro_code": XXXX_model.get("euro_code", ""),
                    "brand": car_model.get("brand", ""),
                    "model": car_model.get("model", ""),
                    "body": car_model.get("body", ""),
                    "year": car_model.get("year", ""),
                    "type_glass": XXXX_model.get("type_glass", ""),
                    "manufacturer": XXXX_model.get("manufacturer", ""),
                    "note": XXXX_model.get("note", ""),
                    "desc": XXXX_model.get("desc", "")
                }
                combined_data.append(combined_model)
                id_counter += 1

    # write the combined data into a new file
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(combined_data, f, indent=2, ensure_ascii=False)

    print(f"Дані були об'єднані і збережені у файлі {output_file}.")


# Delete all json files except the required file
def delete_all_json_files_except(file_to_keep):
    for filename in os.listdir('.'):
        if filename.endswith('.json') and filename != file_to_keep:
            os.remove(filename)


# Options for running the browser in headless mode
options = Options()
options.add_argument('--headless')

# The path to the Chrome driver
driver_path = '/chromedriver'

# You have to download driver for your browser and put it in current directory
# Start a web browser (for example, Chrome)
driver = webdriver.Chrome()

# Open the web page
driver.get("http://XXXX.ua/ukr/login.html")

# find the email and password fields and send the data
email_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "input_login_inpt"))
)
email_input.send_keys(login)

password_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "input_passwd_inpt"))
)
password_input.send_keys(password)

# Click the "Ввійти" button
login_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "#input_sbm input[type='submit']"))
)
login_button.click()
sleep(5)

# Click the "Каталог" button
catalog_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "a.another_li"))
)
catalog_button.click()

# Wait until the "Скачать" (Download) button is clickable
download_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "img[onclick*='/i/documents/Catalog_XXXX.xls']"))
)
# Click the "Скачать" (Download) button
download_button.click()
# Now the file download process should begin.

# Get the file URL from the element's onclick attribute
onclick_attribute = download_button.get_attribute("onclick")
url_start = onclick_attribute.find("'") + 1
url_end = onclick_attribute.rfind("'")
file_url = onclick_attribute[url_start:url_end]

# The full file URL using the current page URL and the resulting file URL
full_file_url = urljoin(driver.current_url, file_url)

# Upload the file to the current directory
response = requests.get(full_file_url)
with open("Catalog_XXXX.xls", "wb") as f:
    f.write(response.content)

# The "Catalog_XXXX.xls" file will now be saved in the current directory
sleep(5)
WebDriverWait(driver, 10).until(EC.url_to_be("http://XXXX.ua/ukr/shop/price.html"))

# Click the "Інтернет магазин" button
shop_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.first_li'))
)
shop_button.click()

sleep(5)
WebDriverWait(driver, 10).until(EC.url_to_be("http://XXXX.ua/ukr/shop/items.html"))

# We open the workbook
workbook = xlrd.open_workbook(file_path)

# Get a list of names of all letters
sheet_names = workbook.sheet_names()

# Select the first sheet
sheet = workbook.sheet_by_name(sheet_names[3])

if not os.path.exists("Car_models.json"):

    # Open the file for recording
    with open("Car_models.json", "w", encoding="utf-8") as f:
        # Display the values of all lines on our letter
        for row_index in range(8, sheet.nrows):
            row_values = sheet.row_values(row_index)

            Car_models = {
                "Model code": f"{row_values[1]}",
                "brand": f"{row_values[2].split(' ')[0]}",
                "model": f"{row_values[2].split(' ')[1]}",
                "body": f"{row_values[3]}",
                "year": f"{row_values[4]}".split('.')[0] + " - " + f"{row_values[5]}".split('.')[0],
            }

            # Write the table data into a file in JSON format
            json.dump(Car_models, f, ensure_ascii=False)
            f.write('\n')

    print("В поточній директорії вже було створено файл 'Car_models.json'")
    sleep(5)

# Create a copy of the "Car_models.json" file with the name "Car_models_full.json"
shutil.copyfile('Car_models.json', 'Car_models_full.json')

# Loading data from the file "Car_models.json"
with open('Car_models.json', 'r', encoding='utf-8') as f:
    car_models_data = [json.loads(line) for line in f]

# Code processing loop
count = 0
for car_model_data in car_models_data:
    # Get the "Model code" from the current car data
    code = car_model_data["Model code"]

    # Find the input field with the identifier "code" and enter the value of the code
    code_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "code"))
    )

    # Clear the input field (if there is a previous value)
    code_input.clear()

    # send the current code to the input field
    code_input.send_keys(code)

    # Wait 12 seconds (save data)
    sleep(12)

    try:
        # receive the HTML code of the page after sending the code
        html = driver.page_source
    except UnexpectedAlertPresentException:
        # If an unexpected pop-up alert appears, close it and continue the iteration
        driver.switch_to.alert.accept()
        continue

    # Create a BeautifulSoup object for HTML parsing
    soup = Bs(html, "html.parser")

    # Find all lines <tr class="found">
    found_rows = soup.find_all("tr", class_="found")

    # Create an empty list to store prices
    glass_price_size_list = []

    if found_rows:
        for row in found_rows:
            # Find the second <td> element in the current line (index 3, since indexes start with 0)
            euro_code_element = row.find_all('td')[0]
            glass_element = row.find_all('td')[1]
            price_element = row.find_all('td')[3]

            # Get the text values of the found elements and add them to the list
            euro_code = euro_code_element.text.split('-')[0].strip()
            glass = glass_element.text.strip()
            price_text = price_element.text.strip()

            # Convert the price_text to a numeric value and add 17%
            price = round(float(price_text) + (float(price_text) * 0.17))

            # Add details to the list
            glass_price_size_list.append(
                {"code": code, "euro_code": euro_code, "name": glass, "price": price,
                 "status": "In stock"})

    else:
        # The code was not found, we are setting the status "Out of stock"
        glass_price_size_list.append({"code": code, "status": "Out of stock"})

    try:
        # Convert the list into JSON format for the current code with the parameter ensure_ascii=False
        json_data = json.dumps(glass_price_size_list, indent=2, ensure_ascii=False)

        # save the JSON data for the current code in the "XXXXXXXX.json" file in the add mode
        with open("XXXX.json", "a", encoding="utf-8") as f:
            f.write(json_data + "\n")

    except Exception as e:
        # Handling an exception if it occurs and outputting an error message
        print(f"Виникла помилка під час запису JSON-даних: {e}")

  # Loading data from the file "Car_models.json"
    with open('Car_models.json', 'r', encoding='utf-8') as f:
        car_models_data = [json.loads(line) for line in f]

    # Delete the first line from the list
    car_models_data.pop(0)

    # Write the processed data back to the "Car_models.json" file
    with open("Car_models.json", "w", encoding="utf-8") as f:
        for car_model_data in car_models_data:
            json.dump(car_model_data, f, ensure_ascii=False)
            f.write('\n')

    # Call the function to convert the files
    format_car_models('Car_models_full.json', 'Car_models_formatted.json')
    process_XXXX_data('XXXX.json', 'XXXX_formatted.json')
    sleep(2)

       # Check whether we processed the last object
    if count == len(car_models_data):
        break
    # Increase the counter
    count += 1
    combine_data('XXXX_formatted.json', 'Car_models_formatted.json', 'Combined_data.json')
# Close the web browser
driver.quit()

# Calling a function to combine data
sleep(5)
# Calling a function to delete all json files except the required file
delete_all_json_files_except('Combined_data.json')
print("Всі дані успішно записано")



