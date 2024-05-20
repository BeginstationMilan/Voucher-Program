import gspread
import time
import random
import string
import img2pdf
import os
import threading
import subprocess
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from gspread.exceptions import SpreadsheetNotFound
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
import win32com.client as win32
import img2pdf
import os
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import win32ui
import win32con
import win32print
import barcode
import base64
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

Code128.default_writer_options['write_text'] = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, 'Output')
DEFAULT_DIR = os.path.join(BASE_DIR, 'Default')


SCOPES = ['https://www.googleapis.com/auth/gmail.send']


if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def loading_bar(action, *args):
    def run_action():
        action(*args)
    
    thread = threading.Thread(target=run_action)
    thread.start()
    
    print("Loading", end="", flush=True)
    for _ in range(40):
        if not thread.is_alive():
            break
        print(".", end="", flush=True)
        time.sleep(0.1)
    print("\n", end="")  # Move to the next line after loading bar
    thread.join()

def generate_unique_code():
    return ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(8))

def generate_barcode(code, filename):
    # Generate barcode and save as PNG
    EAN = barcode.get_barcode_class('ean13')
    ean = EAN(code, writer=ImageWriter())
    barcode_path = f"{filename}.png"
    ean.save(barcode_path)
    return barcode_path

# Define the path to the credentials.json file
CREDENTIALS_FILE_PATH = os.path.join(BASE_DIR, 'credentials.json')

# Modify the get_credentials() function to use the new variable

def get_credentials(credentials_file_path=None):
    if not hasattr(get_credentials, "_creds"):
        try:
            if credentials_file_path is None:
                # If credentials_file_path is not provided, use the directory of the script
                script_dir = os.path.dirname(os.path.abspath(__file__))
                credentials_file_path = os.path.join(script_dir, 'credentials.json')
            
            creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file_path)
            get_credentials._creds = creds
        except Exception as e:
            print(f"Error loading credentials: {str(e)}")
            get_credentials._creds = None
    return get_credentials._creds

def add_to_google_sheet(num_codes, credentials_file_path=CREDENTIALS_FILE_PATH):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Failed to load credentials. Exiting.")
            return

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1

        start_index = sheet.row_count + 1

        # Generate and add new codes
        batch_data = []
        for _ in range(num_codes):
            unique_code = '*' + generate_unique_code()  # Add * to indicate inactive code
            batch_data.append([unique_code])

        sheet.append_rows(batch_data)

        print(f"{num_codes} codes generated and added to the sheet.")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def activate_code(code, credentials_file_path=None):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Failed to load credentials. Exiting.")
            return

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)
        if cell:
            if not cell.value.startswith("*"):  # Check if the code is already activated
                print("Code is already activated.")
            elif sheet.cell(cell.row, 3).value:  # Check if the code has been marked as printed
                sheet.update_cell(cell.row, 1, cell.value[1:])  # Remove the * from the code
                expiration_date = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime("%m/%d/%Y")
                sheet.update_cell(cell.row, 4, expiration_date)  # Update expiration date in column D
                print(f"Code activated successfully. It is valid until {expiration_date}.")
            else:
                print("Code has not been printed. Activation is not allowed.")
        else:
            print("Code not found.")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def check_code_validity(code, credentials_file_path=CREDENTIALS_FILE_PATH):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Failed to load credentials. Exiting.")
            return None

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1

        # Fetch all code data
        data = sheet.get_all_values()

        for row in data:
            if row and (row[0] == code or row[0] == "~" + code or row[0] == "*" + code):
                code_value = row[0]
                if code_value.startswith("~"):
                    redemption_date = row[4]
                    print(f"Code is redeemed on {redemption_date}.")
                elif code_value.startswith("*"):
                    print("Code is not activated.")
                else:
                    expiration_date = row[3]
                    print(f"Code is valid and activated. It is valid until {expiration_date}.")
                    return expiration_date  # Return the expiration date
                break
        else:
            print("Code not found.")
            return None  # Return None when code is not found
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def print_image_in_portrait(file_path):
    try:
        # Convert the image to PDF
        pdf_bytes = img2pdf.convert(file_path)
        
        # Save the PDF file
        pdf_filename = "output.pdf"
        with open(pdf_filename, "wb") as f:
            f.write(pdf_bytes)
        
        # Print the PDF file using the default PDF viewer
        subprocess.Popen([pdf_filename], shell=True).wait()
        
        print("Image has been sent to the printer.")
        
    except Exception as e:
        print(f"An error occurred: {e}")


def save_coupon(code):
    try:
        # Generate barcode and save as PNG
        image_filename = os.path.join(DEFAULT_DIR, 'standaard2.png')  # Updated filename
        barcode_filename = f"barcode_{code}"
        code_without_asterisk = code.replace('*', '')
        code128 = Code128(code_without_asterisk, writer=ImageWriter())
        barcode_img = code128.render()

        if not os.path.exists(image_filename):
            print("Error: Image file not found.")
            return None

        with Image.open(image_filename) as img:
            modified_img = img.copy()
            stretched_img = modified_img.resize((modified_img.width * 2, modified_img.height * 2))
            barcode_img = barcode_img.resize((int(barcode_img.width * 1.7), int(barcode_img.height * 0.7)))  
            position = (35, 675)
            stretched_img.paste(barcode_img, position)

            code_without_asterisk = code.replace('*', '')

            draw = ImageDraw.Draw(stretched_img)
            font = ImageFont.truetype("arial.ttf", 20)
            draw.text((20, 20), f"Code: {code_without_asterisk}", fill="black", font=font)

            output_path = os.path.join(OUTPUT_DIR, f"Coupon_{code_without_asterisk}.png")
            stretched_img.save(output_path)

            return output_path

    except Exception as e:
        print(f"An error occurred while generating coupon PNG: {str(e)}")
        return None


import subprocess

import os

from PIL import ImageDraw, ImageFont

from PIL import ImageDraw, ImageFont

def make_printed():
    try:
        codes = []

        creds = get_credentials()
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        values = sheet.get_all_values()

        # Fetch unprinted codes
        unactivated_codes = [(row[0], i + 1) for i, row in enumerate(values[1:]) if len(row) >= 3 and row[0].startswith("*") and not row[2]]

        if len(unactivated_codes) < 36:
            print("Not enough unactivated codes available.")
            return

        selected_codes = random.sample(unactivated_codes, 36)
        for code, row_index in selected_codes:
            codes.append(code)
            sheet.update_cell(row_index + 1, 3, "Printed")
            print(f"Code {code} marked as printed.")

        final_image_filename = os.path.join(DEFAULT_DIR, 'final.png')
        final_output_path = os.path.join(OUTPUT_DIR, 'final_output.png')
        coupon_list_pdf = os.path.join(OUTPUT_DIR, 'Coupon_List.pdf')

        if not os.path.exists(final_image_filename):
            print("Error: Final image file not found.")
            return

        with Image.open(final_image_filename) as final_img:
            draw = ImageDraw.Draw(final_img)
            font = ImageFont.load_default()

            coupon_width, coupon_height = 204, 325
            positions = [(x * coupon_width, y * coupon_height) for y in range(6) for x in range(6)]

            for code in codes:
                image_filename = os.path.join(DEFAULT_DIR, 'standaard.png')
                if not os.path.exists(image_filename):
                    print("Error: Image file not found.")
                    return

                with Image.open(image_filename) as img:
                    code_without_asterisk = code.replace('*', '')
                    code128 = Code128(code_without_asterisk, writer=ImageWriter())
                    barcode_img = code128.render()

                    barcode_width, barcode_height = 300, 75
                    barcode_x, barcode_y = 10, 350

                    barcode_img = barcode_img.resize((barcode_width, barcode_height))
                    img.paste(barcode_img, (barcode_x, barcode_y))

                    coupon_img = img.resize((coupon_width, coupon_height))

                    pos_x, pos_y = positions[codes.index(code)]
                    final_img.paste(coupon_img, (pos_x, pos_y))

                    draw.text((pos_x + 10, pos_y + 10), f"Code: {code_without_asterisk}", fill="black", font=font)

            final_img.save(final_output_path)
            print("Output image combined with final design successfully.")

            images = [final_output_path, os.path.join(DEFAULT_DIR, 'standaard_full.png')]
            pdf_bytes = img2pdf.convert(images)

            with open(coupon_list_pdf, "wb") as f:
                f.write(pdf_bytes)

            print("Coupon List PDF generated successfully.")
            os.startfile(coupon_list_pdf)

    except Exception as e:
        print(f"An error occurred: {str(e)}")

def find_random_unprinted_code(credentials_file_path=None):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Failed to load credentials. Exiting.")
            return None

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        all_codes = sheet.col_values(1)[1:]
        unprinted_codes = [code for code in all_codes if not sheet.cell(all_codes.index(code) + 1, 3).value]
        if unprinted_codes:
            return random.choice(unprinted_codes)
        else:
            print("No unprinted codes found.")
            return None
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
        return None
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None
    
def mark_as_printed(code, credentials_file_path=CREDENTIALS_FILE_PATH):
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file_path)
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)
        if cell:
            sheet.update_cell(cell.row, 3, "Printed")  # Update column C with "Printed"
            print(f"Code {code} marked as printed.")
        else:
            print("Code not found.")
    except SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Function to print image in Chrome
def print_images_in_windows(file_paths):
    try:
        # Create a printer device context
        printer_name = win32print.GetDefaultPrinter()
        printer_dc = win32print.OpenPrinter(printer_name)
        printer_handle = win32print.GetPrinter(printer_dc, 2)['hPrinter']
        printer_dc = win32ui.CreateDC()
        printer_dc.CreatePrinterDC(printer_name)

        for file_path in file_paths:
            # Open the image file
            image = Image.open(file_path)

            # Get the dimensions of the image
            width, height = image.size

            # Start a print job
            printer_dc.StartDoc(file_path)
            printer_dc.StartPage()

            # Draw the image on the printer device context
            dib = ImageWin.Dib(image)
            dib.draw(printer_dc.GetHandleOutput(), (0, 0, width, height))

            # End the print job
            printer_dc.EndPage()
            printer_dc.EndDoc()

        # Close the printer device context
        printer_dc.Close()

        print("Images have been sent to the printer.")

    except Exception as e:
        print(f"An error occurred: {e}")


def info_status():
    while True:
        clear_screen()
        print("Info Status Menu:")
        print("1. Basic Info")
        print("2. All Codes")
        print("3. Return to Main Menu")

        choice = input("Enter your choice: ")

        if choice == '1':
            show_basic_info()
        elif choice == '2':
            show_all_codes()
        elif choice == '3':
            print("Returning to Main Menu...")
            return
        else:
            print("Invalid choice. Please try again.")

def show_basic_info():
    try:
        creds = get_credentials()
        if creds is None:
            print("Failed to load credentials. Exiting.")
            return
        
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1

        # Retrieve all data in one request
        data = sheet.get_all_values()

        total_codes = len(data) - 1  # Exclude header row
        active_codes = sum(1 for row in data[1:] if row and not row[0].startswith("*") and not row[0].startswith("~"))
        inactive_codes = sum(1 for row in data[1:] if row and row[0].startswith("*"))
        printed_codes = sum(1 for row in data[1:] if row and row[2])
        used_codes = sum(1 for row in data[1:] if row and row[0].startswith("~"))

        print("Basic Info:")
        print(f"Total codes: {total_codes}")
        print(f"Active codes: {active_codes}")
        print(f"Inactive codes: {inactive_codes}")
        print(f"Printed codes: {printed_codes}")
        print(f"Used codes: {used_codes}")
        print("Program created by: Milan Vosters")
        print("Project started on: 08:00AM 16/5/2024")
        input("Press any key to return to the Info Status Menu...")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        input("Press any key to return to the Info Status Menu...")

def show_all_codes():
    try:
        creds = get_credentials()
        if creds is None:
            print("Failed to load credentials. Exiting.")
            return
        
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1

        # Retrieve all data in one request
        data = sheet.get_all_values()

        print("All Codes:")
        for row in data[1:]:
            if row:
                code = row[0]
                code_status = "Active" if code and not code.startswith("*") and not code.startswith("~") else "Inactive" if code and code.startswith("*") else "Used" if code and code.startswith("~") else "Unknown"
                print(f"Code: {code}, Status: {code_status}")

        input("Press any key to return to the Info Status Menu...")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        input("Press any key to return to the Info Status Menu...")



def redeem_code(code):
    try:
        creds = get_credentials()
        if creds is None:
            print("Failed to load credentials. Exiting.")
            return

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)
        if cell:
            if cell.value.startswith("*"):  # Check if the code is inactive
                print("Code is inactive and cannot be redeemed.")
            elif not cell.value.startswith("~"):  # Check if the code is already redeemed
                sheet.update_cell(cell.row, 1, "~" + cell.value)  # Add ~ to the code
                sheet.update_cell(cell.row, 5, datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S"))  # Add redemption time to column E
                print("Code redeemed successfully.")
            else:
                print("Code is already redeemed.")
        else:
            print("Code not found.")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


import win32com.client as win32

def main():
    """
    The main function that provides a menu to the user and handles user input.
    """
    while True:
        clear_screen()
        print("\nMenu:")
        print("1. Generate codes")
        print("2. Check code validity")
        print("3. Print")
        print("4. Info status")
        print("5. Activate code")
        print("6. Redeem code")
        print("7. Exit")

        choice = input("Enter your choice: ")

        if choice == '1':
            clear_screen()
            num_codes = int(input("How many codes do you want to generate? "))
            clear_screen()
            add_to_google_sheet(num_codes)
        elif choice == '2':
            clear_screen()
            code = input("Enter code to check validity: ")
            clear_screen()
            check_code_validity(code)
        elif choice == '3':
            clear_screen()
            make_printed()
        elif choice == '4':
            clear_screen()
            info_status()
        elif choice == '5':
            clear_screen()
            code = input("Enter code to activate: ")
            clear_screen()
            activate_code(code)
        elif choice == '6':
            clear_screen()
            code = input("Enter code to redeem: ")
            clear_screen()
            redeem_code(code)
        elif choice == '7':
            print("Exiting program...")
            break
        else:
            print("Invalid choice. Please try again.")
            input("Press Enter to continue...")  # Wait for user input before returning to the main menu
        
        time.sleep(3)

    print("End of program.")

    # Call the main function to start the program
if __name__ == "__main__":
    main()