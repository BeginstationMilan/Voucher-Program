import gspread
import time
import sys
import threading
import random
import colorama
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
from alive_progress import alive_bar
import warnings

warnings.simplefilter("ignore", SyntaxWarning)

Code128.default_writer_options['write_text'] = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, 'Output')
DEFAULT_DIR = os.path.join(BASE_DIR, 'Default')


SCOPES = ['https://www.googleapis.com/auth/gmail.send']


if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')
    
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
            print(f"Error bij het laden vam credentials: {str(e)}")
            get_credentials._creds = None
    return get_credentials._creds

def add_to_google_sheet(num_codes, credentials_file_path=CREDENTIALS_FILE_PATH):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Gefaald om credentials te laden. Sluiten.")
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

        print(f"{num_codes} codes gegenereed en in spreadsheet geplaatst.")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' niet gevonden.")
    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")

def activate_code(code, credentials_file_path=None):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Gefaald om credentials te laden. Sluiten.")
            return

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)
        if cell:
            if not cell.value.startswith("*"):  # Check if the code is already activated
                print("Code is al geactiveerd.")
            elif sheet.cell(cell.row, 3).value:  # Check if the code has been marked as printed
                sheet.update_cell(cell.row, 1, cell.value[1:])  # Remove the * from the code
                expiration_date = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime("%m/%d/%Y")
                sheet.update_cell(cell.row, 4, expiration_date)  # Update expiration date in column D
                print(f"Code geactiveerd. is geldig tot {expiration_date}.")
            else:
                print("Code is niet geprint. Activatie is niet mogelijk.")
        else:
            print("Code niet gevonden.")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' niet gevonden.")
    except Exception as e:
        print(f"An error is gebeurd: {str(e)}")

def check_code_validity(code, credentials_file_path=CREDENTIALS_FILE_PATH):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Gefaald om credentials te laden. Sluiten.")
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
                    print(f"Code is ingeleverd op {redemption_date}.")
                    time.sleep(3)
                elif code_value.startswith("*"):
                    print("Code is niet geactiveerd.")
                else:
                    expiration_date = row[3]
                    print(f"Code is geldig en geactiveerd. is geldig tot {expiration_date}.")
                    return expiration_date  # Return the expiration date
                break
        else:
            print("Code niet gevonden.")
            return None  # Return None when code is not found
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' niet gevonden.")
    except Exception as e:
        print(f"een error is gebeurd: {str(e)}")

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
        
        print("Foto is gestuurd naar printer.")
        
    except Exception as e:
        print(f"een error is gebeurd: {e}")


def save_coupon(code):
    try:
        # Generate barcode and save as PNG
        image_filename = os.path.join(DEFAULT_DIR, 'standaard2.png')  # Updated filename
        barcode_filename = f"barcode_{code}"
        code_without_asterisk = code.replace('*', '')
        code128 = Code128(code_without_asterisk, writer=ImageWriter())
        barcode_img = code128.render()

        if not os.path.exists(image_filename):
            print("Error: Foto niet gevonden.")
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
        print(f"Een error is gebeurd tijdens het genereren van Coupon: {str(e)}")
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
            print("Niet genoeg onactieve codes beschikbaar.")
            return

        selected_codes = random.sample(unactivated_codes, 36)
        for code, row_index in selected_codes:
            codes.append(code)
            sheet.update_cell(row_index + 1, 3, "Geprint")
            print(f"Code {code} Gevonden.")

        final_image_filename = os.path.join(DEFAULT_DIR, 'final.png')
        final_output_path = os.path.join(OUTPUT_DIR, 'final_output.png')
        coupon_list_pdf = os.path.join(OUTPUT_DIR, 'Coupon_List.pdf')

        if not os.path.exists(final_image_filename):
            print("Error: Final foto bestand is niet gevonden.")
            return

        with Image.open(final_image_filename) as final_img:
            draw = ImageDraw.Draw(final_img)
            font = ImageFont.load_default()

            coupon_width, coupon_height = 204, 325
            positions = [(x * coupon_width, y * coupon_height) for y in range(6) for x in range(6)]

            for code in codes:
                image_filename = os.path.join(DEFAULT_DIR, 'standaard.png')
                if not os.path.exists(image_filename):
                    print("Error: Foto bestand niet gevonden.")
                    return

                with Image.open(image_filename) as img:
                    code_without_asterisk = code.replace('*', '')
                    code128 = Code128(code_without_asterisk, writer=ImageWriter())
                    barcode_img = code128.render()

                    barcode_width, barcode_height = 340, 75
                    barcode_x, barcode_y = -10, 350

                    barcode_img = barcode_img.resize((barcode_width, barcode_height))
                    img.paste(barcode_img, (barcode_x, barcode_y))

                    coupon_img = img.resize((coupon_width, coupon_height))

                    pos_x, pos_y = positions[codes.index(code)]
                    final_img.paste(coupon_img, (pos_x, pos_y))

                    draw.text((pos_x + 10, pos_y + 10), f"Code: {code_without_asterisk}", fill="black", font=font)

            final_img.save(final_output_path)
            print("\nPagina Genereren.........")
            time.sleep(2)
            print("\nOutput image gecombineerd met final design.")

            images = [final_output_path, os.path.join(DEFAULT_DIR, 'standaard_full.png')]
            pdf_bytes = img2pdf.convert(images)

            with open(coupon_list_pdf, "wb") as f:
                f.write(pdf_bytes)

            print("Coupon Lijst PDF gegenereerd.")
            os.startfile(coupon_list_pdf)

    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")

def find_random_unprinted_code(credentials_file_path=None):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Gefaald om credentials te laden. Sluiten.")
            return None

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        all_codes = sheet.col_values(1)[1:]
        unprinted_codes = [code for code in all_codes if not sheet.cell(all_codes.index(code) + 1, 3).value]
        if unprinted_codes:
            return random.choice(unprinted_codes)
        else:
            print("Geen ongeprinten codes gevonden.")
            return None
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' niet gevonden.")
        return None
    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")
        return None
    
def mark_as_printed(code, credentials_file_path=CREDENTIALS_FILE_PATH):
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file_path)
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)
        if cell:
            sheet.update_cell(cell.row, 3, "Geprint")  # Update column C with "Printed"
            print(f"Code {code} gemarkeerd als geprint.")
        else:
            print("Code niet gevonden.")
    except SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' niet gevonden.")
    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")

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

        print("Fotos zijn gestuurd naar de printer.")

    except Exception as e:
        print(f"Een error is gebeurd: {e}")


def info_status():
    while True:
        clear_screen()
        print("\033[1m   _____        __         \033[0m")
        print("\033[1m  |_   _|      / _|        \033[0m")
        print("\033[1m    | |  _ __ | |_ ___     \033[0m")
        print("\033[1m    | | | '_ \|  _/ _ \    \033[0m")
        print("\033[1m   _| |_| | | | || (_) |   \033[0m")
        print("\033[1m  |_____|_| |_|_| \___/    \033[0m")
        print("\033[1m  |  \/  |                 \033[0m")
        print("\033[1m  | \  / | ___ _ __  _   _ \033[0m")
        print("\033[1m  | |\/| |/ _ \ '_ \| | | |\033[0m")
        print("\033[1m  | |  | |  __/ | | | |_| |\033[0m")
        print("\033[1m  |_|  |_|\___|_| |_|\__,_|\033[0m")
        print("\033[1m                           \033[0m")
        print("\033[1m                           \033[0m")
        print("\nInfo Status Menu:")
        print("1. Basis Info")
        print("2. Alle Codes")
        print("3. Keer terug naar Main Menu")

        choice = input("Toets jouw keuze in: ")

        if choice == '1':
            show_basic_info()
        elif choice == '2':
            show_all_codes()
        elif choice == '3':
            print("Aan het terug keren naar Main Menu...")
            return
        else:
            print("Ongeldige keuze. Probeer opnieuw.")

def show_basic_info():
    try:
        creds = get_credentials()
        if creds is None:
            print("Gefaald om credentials te laden. Sluiten.")
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

        print("Basis Info:")
        print(f"Totaal aantal codes: {total_codes}")
        print(f"Actieve codes: {active_codes}")
        print(f"Inactieve codes: {inactive_codes}")
        print(f"Geprinte codes: {printed_codes}")
        print(f"Gebruikte codes: {used_codes}")
        print("Programma gemaakt door: Milan Vosters")
        print("Project gestart op: 08:00AM 16/5/2024")
        input("Klik een toets om terug te keren naar het info menu...")
    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")
        input("Klik een toets om terug te keren naar het info menu...")

def show_all_codes():
    try:
        creds = get_credentials()
        if creds is None:
            print("Gefaald om credentials te laden. Sluiten.")
            return
        
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1

        # Retrieve all data in one request
        data = sheet.get_all_values()

        print("Alle Codes:")
        for row in data[1:]:
            if row:
                code = row[0]
                code_status = "Actieve" if code and not code.startswith("*") and not code.startswith("~") else "Inactief" if code and code.startswith("*") else "Gebruikt" if code and code.startswith("~") else "Unknown"
                print(f"Code: {code}, Status: {code_status}")
                time.sleep(0.03)

        input("Klik een toets om terug te keren naar het status menu...")
    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")
        input("Klik een toets om terug te keren naar het status menu...")



def redeem_code(code):
    try:
        creds = get_credentials()
        if creds is None:
            print("Gefaald om credentials te laden. Sluiten.")
            return

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)
        if cell:
            if cell.value.startswith("*"):  # Check if the code is inactive
                print("Code is inactief en kan niet worden ingeleverd.")
            elif not cell.value.startswith("~"):  # Check if the code is already redeemed
                sheet.update_cell(cell.row, 1, "~" + cell.value)  # Add ~ to the code
                sheet.update_cell(cell.row, 5, datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S"))  # Add redemption time to column E
                print("Code ingeleverd.")
            else:
                print("Code is al ingeleverd.")
        else:
            print("Code niet gevonden.")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' niet gevonden.")
    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")


import win32com.client as win32

def main():
    """
    The main function that provides a menu to the user and handles user input.
    """
    while True:
        clear_screen()
        print("\033[1m __      __              _                  \033[0m")
        print("\033[1m \ \    / /             | |                 \033[0m")
        print("\033[1m  \ \  / /__  _   _  ___| |__   ___ _ __    \033[0m")
        print("\033[1m   \ \/ / _ \| | | |/ __| '_ \ / _ \ '__|   \033[0m")
        print("\033[1m    \  / (_) | |_| | (__| | | |  __/ |      \033[0m")
        print("\033[1m  ___\/ \___/ \__,_|\___|_| |_|\___|_|      \033[0m")
        print("\033[1m |  __ \                                     \033[0m")
        print("\033[1m | |__) | __ ___   __ _ _ __ __ _ _ __ ___  \033[0m")
        print("\033[1m |  ___/ '__/ _ \ / _` | '__/ _` | '_ ` _ \ \033[0m")
        print("\033[1m | |   | | | (_) | (_| | | | (_| | | | | | |\033[0m")
        print("\033[1m |_|   |_|  \___/ \__, |_|  \__,_|_| |_| |_|\033[0m")
        print("\033[1m                   __/ |                    \033[0m")
        print("\033[1m                  |___/                     \033[0m")
        print("\nDit Programma was gemaakt door: Milan Vosters")

        print("\nMenu:")
        print("\n 1. Genereer codes")
        print(" 2. Verifieer code")
        print(" 3. Print codes")
        print(" 4. Info status")
        print(" 5. Activeer code")
        print(" 6. Lever code in")
        print(" 7. Sluit af")

        choice = input("\nKies je keuze: ")
        
        if choice == '1':
            clear_screen()
            print("\033[1m   _____          _                              \033[0m")
            print("\033[1m  / ____|        | |                             \033[0m")
            print("\033[1m | |     ___   __| | ___                         \033[0m")
            print("\033[1m | |    / _ \ / _` |/ _ \                        \033[0m")
            print("\033[1m | |___| (_) | (_| |  __/                        \033[0m")
            print("\033[1m  \_____\___/ \__,_|\___|          _             \033[0m")
            print("\033[1m  / ____|                         | |            \033[0m")
            print("\033[1m | |  __  ___ _ __   ___ _ __ __ _| |_ ___  _ __ \033[0m")
            print("\033[1m | | |_ |/ _ \ '_ \ / _ \ '__/ _` | __/ _ \| '__|\033[0m")
            print("\033[1m | |__| |  __/ | | |  __/ | | (_| | || (_) | |   \033[0m")
            print("\033[1m  \_____|\___|_| |_|\___|_|  \__,_|\__\___/|_|   \033[0m")
            print("\033[1m                                                 \033[0m")
            print("\033[1m                                                 \033[0m")
            num_codes = int(input("\nHoeveel codes wil je genereren? "))
            add_to_google_sheet(num_codes)
        elif choice == '2':
            clear_screen()
            print("\033[1m  __      __       _  __       \033[0m")
            print("\033[1m  \ \    / /      (_)/ _|      \033[0m")
            print("\033[1m   \ \  / /__ _ __ _| |_ _   _ \033[0m")
            print("\033[1m    \ \/ / _ \ '__| |  _| | | |\033[0m")
            print("\033[1m     \  /  __/ |  | | | | |_| |\033[0m")
            print("\033[1m    __\/_\___|_|  |_|_|  \__, |\033[0m")
            print("\033[1m   / ____|        | |     __/ |\033[0m")
            print("\033[1m  | |     ___   __| | ___|___/ \033[0m")
            print("\033[1m  | |    / _ \ / _` |/ _ \     \033[0m")
            print("\033[1m  | |___| (_) | (_| |  __/     \033[0m")
            print("\033[1m   \_____\___/ \__,_|\___|     \033[0m")
            print("\033[1m                               \033[0m")
            print("\033[1m                               \033[0m")
            code = input("\nvul code in om te verifieren: ")
            check_code_validity(code)
        elif choice == '3':
            clear_screen()
            print("\033[1m   _____      _       _        \033[0m")
            print("\033[1m  |  __ \    (_)     | |       \033[0m")
            print("\033[1m  | |__) | __ _ _ __ | |_      \033[0m")
            print("\033[1m  |  ___/ '__| | '_ \| __|     \033[0m")
            print("\033[1m  | |   | |  | | | | | |_      \033[0m")
            print("\033[1m  |_|___|_|  |_|_| |_|\__|     \033[0m")
            print("\033[1m   / ____|        | |          \033[0m")
            print("\033[1m  | |     ___   __| | ___  ___ \033[0m")
            print("\033[1m  | |    / _ \ / _` |/ _ \/ __|\033[0m")
            print("\033[1m  | |___| (_) | (_| |  __/\__ \\\033[0m")
            print("\033[1m   \_____\___/ \__,_|\___||___/\033[0m")
            print("\033[1m                               \033[0m")
            print("\033[1m                               \033[0m")
            make_printed()
        elif choice == '4':
            clear_screen()
            info_status()
        elif choice == '5':
            clear_screen()
            print("\033[1m               _   _            _        \033[0m")
            print("\033[1m     /\       | | (_)          | |       \033[0m")
            print("\033[1m    /  \   ___| |_ ___   ____ _| |_ ___  \033[0m")
            print("\033[1m   / /\ \ / __| __| \ \ / / _` | __/ _ \ \033[0m")
            print("\033[1m  / ____ \ (__| |_| |\ V / (_| | ||  __/ \033[0m")
            print("\033[1m /_/____\_\___|\__|_| \_/ \__,_|\__\___| \033[0m")
            print("\033[1m  / ____|        | |                      \033[0m")
            print("\033[1m | |     ___   __| | ___                  \033[0m")
            print("\033[1m | |    / _ \ / _` |/ _ \                 \033[0m")
            print("\033[1m | |___| (_) | (_| |  __/                 \033[0m")
            print("\033[1m  \_____\___/ \__,_|\___|                 \033[0m")
            print("\033[1m                                          \033[0m")
            print("\033[1m                                          \033[0m")
            code = input("\nvul code in om te activeren: ")
            activate_code(code)
        elif choice == '6':
            clear_screen()
            print("\033[1m  _____          _                     \033[0m")
            print("\033[1m |  __ \        | |                    \033[0m")
            print("\033[1m | |__) |___  __| | ___  ___ _ __ ___  \033[0m")
            print("\033[1m |  _  // _ \/ _` |/ _ \/ _ \ '_ ` _ \ \033[0m")
            print("\033[1m | | \ \  __/ (_| |  __/  __/ | | | | |\033[0m")
            print("\033[1m |_|__\_\___|\__,_|\___|\___|_| |_| |_|\033[0m")
            print("\033[1m  / ____|        | |                    \033[0m")
            print("\033[1m | |     ___   __| | ___                \033[0m")
            print("\033[1m | |    / _ \ / _` |/ _ \               \033[0m")
            print("\033[1m | |___| (_) | (_| |  __/               \033[0m")
            print("\033[1m  \_____\___/ \__,_|\___|               \033[0m")
            print("\033[1m                                        \033[0m")
            print("\033[1m                                        \033[0m")
            code = input("\nVul code in om in te leveren: ")
            redeem_code(code)
        elif choice == '7':
            print("Programma afsluiten...")
            break
        else:
            print("Ongeldige keuze. Probeer opnieuw.")
            input("Druk enter om door te gaan...")  # Wait for user input before returning to the main menu
        
        time.sleep(3)

    print("Einde programma.")

if __name__ == "__main__":
    main()