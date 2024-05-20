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
from PIL import Image, ImageOps  # Add this line for the new imports
import win32com.client as win32  
import win32print
import img2pdf
import os
from PIL import Image, ImageDraw, ImageFont
from barcode import Code128
from barcode.writer import ImageWriter
import img2pdf
import win32ui
import win32con
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image, ImageOps
from barcode import Code128
from barcode.writer import ImageWriter
import os
from PIL import ImageDraw, ImageFont

Code128.default_writer_options['write_text'] = False

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
    barcode_folder = r'C:\Users\milan\Desktop\Project Voucher\Output'  # Specify the folder where barcodes will be saved
    if not os.path.exists(barcode_folder):
        os.makedirs(barcode_folder)
    filepath = os.path.join(barcode_folder, filename)
    code128 = Code128(code, writer=ImageWriter())
    code128.save(filepath)
    return filepath


def get_credentials():
    # List of credential files to try
    credential_files = ['credentials.json', 'credentials2.json']  # Add more files if needed

    for cred_file in credential_files:
        try:
            # Attempt to load credentials from the current file
            scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, scope)
            return creds
        except Exception as e:
            print(f"Failed to load credentials from {cred_file}: {str(e)}")

    # None of the credential files were valid
    return None

def add_to_google_sheet(num_codes):
    creds = get_credentials()
    if creds is None:
        print("Failed to load credentials. Exiting.")
        return

    client = gspread.authorize(creds)

    try:
        sheet = client.open('voucher-barcode1').sheet1
        start_index = len(sheet.col_values(1)) + 1

        for i in range(num_codes):
            unique_code = '*' + generate_unique_code()  # Add * to indicate inactive code
            sheet.update_cell(start_index + i, 1, unique_code)
        print(f"{num_codes} codes generated and added to the sheet.")
    except SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def check_code_validity(code):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)

    try:
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)
        if cell:
            code_value = cell.value
            if code_value.startswith("~"):
                redemption_date = sheet.cell(cell.row, 5).value
                print(f"Code is redeemed on {redemption_date}.")
            elif code_value.startswith("*"):
                print("Code is not activated.")
            else:
                expiration_date = sheet.cell(cell.row, 4).value
                print(f"Code is valid and activated. It is valid until {expiration_date}.")
                return expiration_date  # Return the expiration date
        else:
            print("Code not found.")
            return None  # Return None when code is not found
    except SpreadsheetNotFound:
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
        
        # Print the PDF file
        acrobat_path = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"  # Modify this path as per your installation
        subprocess.Popen([acrobat_path, "/t", pdf_filename, win32print.GetDefaultPrinter()], shell=True).wait()
        
        print("Image has been sent to the printer.")
        
    except Exception as e:
        print(f"An error occurred: {e}")

def save_as_png(code):
    try:
        image_filename = r'C:\Users\milan\Desktop\Project Voucher\Default\standaard2.png'  # Updated filename
        barcode_filename = f"barcode_{code}"  # Remove .png extension here

        # Generate the barcode directly without using the generate_barcode function
        code_without_asterisk = code.replace('*', '')  # Remove asterisks from code if present
        code128 = Code128(code_without_asterisk, writer=ImageWriter())
        barcode_img = code128.render()

        # Ensure the image exists before attempting to combine it with the barcode
        if not os.path.exists(image_filename):
            print("Error: Image file not found.")
            return

        with Image.open(image_filename) as img:
            # Create a copy of the original image to avoid modifying it
            modified_img = img.copy()

            # Stretch the copied image to twice its size and twice its length
            stretched_img = modified_img.resize((modified_img.width * 2, modified_img.height * 2))

            # Resize and position the barcode image
            barcode_img = barcode_img.resize((int(barcode_img.width * 1.7), barcode_img.height - 30))
            position = (0, 675)
            stretched_img.paste(barcode_img, position)

            # Remove the * from the code in the filename
            code_without_asterisk = code.replace('*', '')

            # Draw coupon code text on the image
            draw = ImageDraw.Draw(stretched_img)
            font = ImageFont.load_default()
            draw.text((20, 20), f"Code: {code_without_asterisk}", fill="black", font=font)

            # Save the stretched image with the barcode
            output_path = os.path.join(r'C:\Users\milan\Desktop\Project Voucher\Output', f"Coupon_{code_without_asterisk}.png")  # Include code in filename
            stretched_img.save(output_path)

            print("Barcode added to design successfully and saved as PNG.")

            # Open the generated image
            os.startfile(output_path)

    except Exception as e:
        print(f"An error occurred: {str(e)}")

import subprocess

import os

from PIL import ImageDraw, ImageFont

from PIL import ImageDraw, ImageFont

def make_printed():
    try:
        codes = []

        # Fetch all data from the Google Sheet
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        values = sheet.get_all_values()

        # Extract unprinted codes and their row indices
        unprinted_codes = []
        for i, row in enumerate(values[1:], start=2):  # Skip the header row
            if len(row) >= 3 and not row[2]:  # Check if column C is empty (unprinted)
                unprinted_codes.append((row[0], i))  # Store code and its row index

        if len(unprinted_codes) < 9:
            print("Not enough unprinted codes available.")
            return

        # Randomly select 9 unprinted codes
        selected_codes = random.sample(unprinted_codes, 9)

        # Mark selected codes as printed and add them to the list
        for code, row_index in selected_codes:
            codes.append(code)
            sheet.update_cell(row_index, 3, "Printed")  # Update column C
            print(f"Code {code} marked as printed.")

        # Generate the printed coupons
        final_image_filename = r'C:\Users\milan\Desktop\Project Voucher\Default\final.png'
        final_output_path = os.path.join(r'C:\Users\milan\Desktop\Project Voucher\Output', f"final_output.png")
        coupon_list_pdf = os.path.join(r'C:\Users\milan\Desktop\Project Voucher\Output', f"Coupon_List.pdf")

        if not os.path.exists(final_image_filename):
            print("Error: Final image file not found.")
            return

        with Image.open(final_image_filename) as final_img:
            draw = ImageDraw.Draw(final_img)
            font = ImageFont.load_default()

            for code in codes:
                image_filename = r'C:\Users\milan\Desktop\Project Voucher\Default\standaard.png'

                if not os.path.exists(image_filename):
                    print("Error: Image file not found.")
                    return

                coupon_width = 408
                coupon_height = 650
                positions = [
                    (0, 0), (408, 0), (816, 0),
                    (0, 650), (408, 650), (816, 650),
                    (0, 1300), (408, 1300), (816, 1300)
                ]

                with Image.open(image_filename) as img:
                    code_without_asterisk = code.replace('*', '')  # Remove the asterisks from the code

                    code128 = Code128(code_without_asterisk, writer=ImageWriter())
                    barcode_img = code128.render()

                    barcode_width = 300
                    barcode_height = 75
                    barcode_x = 15
                    barcode_y = 350

                    barcode_img = barcode_img.resize((barcode_width, barcode_height))
                    img.paste(barcode_img, (barcode_x, barcode_y))

                    coupon_img = img.resize((coupon_width, coupon_height))
                    pos_x, pos_y = positions[codes.index(code)]
                    final_img.paste(coupon_img, (pos_x, pos_y))

                    # Draw coupon code text on the image
                    draw.text((pos_x + 10, pos_y + 10), f"Code: {code_without_asterisk}", fill="black", font=font)

            final_img.save(final_output_path)
            print("Output image combined with final design successfully.")

            # Convert the images to PDF
            images = [final_output_path, r'C:\Users\milan\Desktop\Project Voucher\Default\standaard_full.png']
            pdf_bytes = img2pdf.convert(images)

            # Save the PDF file
            with open(coupon_list_pdf, "wb") as f:
                f.write(pdf_bytes)

            print("Coupon List PDF generated successfully.")

            # Open the generated image
            os.startfile(coupon_list_pdf)

    except Exception as e:
        print(f"An error occurred: {str(e)}")

def find_random_unprinted_code():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)

    try:
        sheet = client.open('voucher-barcode1').sheet1
        all_codes = sheet.col_values(1)[1:]
        unprinted_codes = [code for code in all_codes if not sheet.cell(all_codes.index(code) + 1, 3).value]
        if unprinted_codes:
            return random.choice(unprinted_codes)
        else:
            print("No unprinted codes found.")
            return None
    except SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
        return None
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

def mark_as_printed(code):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)

    try:
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
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1

        # Retrieve all data in one request
        data = sheet.get_all_values()

        total_codes = len(data) - 1  # Exclude header row
        active_codes = sum(1 for row in data[1:] if row[0] and not row[0].startswith("*") and not row[0].startswith("~"))
        inactive_codes = sum(1 for row in data[1:] if row[0] and row[0].startswith("*"))
        printed_codes = sum(1 for row in data[1:] if row[0] and row[2])
        used_codes = sum(1 for row in data[1:] if row[0] and row[0].startswith("~"))

        print("Basic Info:")
        print(f"Total codes: {total_codes}")
        print(f"Active codes: {active_codes}")
        print(f"Inactive codes: {inactive_codes}")
        print(f"Printed codes: {printed_codes}")
        print(f"Used codes: {used_codes}")
        print("Program created by: Milan Vosters")
        print("Project started on: 08:00AM 16/5/2024")
        input("Press any key to return to the Info Status Menu...")
    except SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        input("Press any key to return to the Info Status Menu...")

def show_all_codes():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1

        # Retrieve all data in one request
        data = sheet.get_all_values()

        print("All Codes:")
        for row in data[1:]:
            code = row[0]
            code_status = "Active" if code and not code.startswith("*") and not code.startswith("~") else "Inactive" if code and code.startswith("*") else "Used" if code and code.startswith("~") else "Unknown"
            print(f"Code: {code}, Status: {code_status}")

        input("Press any key to return to the Info Status Menu...")
    except SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
        input("Press any key to return to the Info Status Menu...")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        input("Press any key to return to the Info Status Menu...")

def activate_code(code):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)

    try:
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)
        if cell:
            if cell.value.startswith("*"):
                sheet.update_cell(cell.row, 1, cell.value[1:])  # Remove the * from the code
                expiration_date = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime("%m/%d/%Y")
                sheet.update_cell(cell.row, 4, expiration_date)  # Update expiration date in column D
                print(f"Code activated successfully. It is valid until {expiration_date}.")
            else:
                print("Code is already activated.")
        else:
            print("Code not found.")
    except SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def redeem_code(code):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)

    try:
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)
        if cell:
            if not cell.value.startswith("~"):
                sheet.update_cell(cell.row, 1, "~" + cell.value)  # Add ~ to the code
                sheet.update_cell(cell.row, 5, datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S"))  # Add redemption time to column E
                print("Code redeemed successfully.")
            else:
                print("Code is already redeemed.")
        else:
            print("Code not found.")
    except SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def activate_and_email(name, email_address):
    """
    Generate a voucher code, activate it, combine with an image, and send it via email.
    """
    try:
        # Generate a random unactivated code
        code = get_random_unactivated_code()
        if code is None:
            print("Not enough codes! Generating more now...")
            add_to_google_sheet(9)  # Generate nine more codes
            code = get_ranzdom_unactivated_code()  # Try finding a code again
            if code is None:
                print("Failed to generate more codes. Exiting.")
                return

        # Activate the code
        activate_code(code)

        # Generate barcode and save it
        code_without_asterisk = code.replace('*', '')  # Remove asterisks from the code
        barcode_filename = f"barcode_{code_without_asterisk}"  # Remove asterisks from the filename
        generate_barcode(code_without_asterisk, barcode_filename)

        # Add a short delay to ensure the barcode generation completes
        time.sleep(1)

        # File paths
        image_filename = r'C:\Users\milan\Desktop\Project Voucher\Default\standaard2.png'
        output_folder = r'C:\Users\milan\Desktop\Project Voucher\Output'
        output_path = os.path.join(output_folder, f'Coupon_email_{code_without_asterisk}.png')  # Use code without asterisks in filename

        # Ensure the image and barcode exist before attempting to combine them
        if not os.path.exists(image_filename):
            print("Error: Image file not found.")
            return
        barcode_filepath = os.path.join(output_folder, f"{barcode_filename}.png")
        if not os.path.exists(barcode_filepath):
            print("Error: Barcode file not found.")
            return

        # Open and modify the image
        with Image.open(image_filename) as img:
            with Image.open(barcode_filepath) as barcode_img:
                # Stretch the image
                modified_img = img.resize((img.width * 2, img.height * 2))

                # Adjust barcode size
                new_barcode_width = int(barcode_img.width * 1.7)
                resized_barcode = barcode_img.resize((new_barcode_width, barcode_img.height - 30))

                # Position to place the barcode image
                position = (20, 675)
                modified_img.paste(resized_barcode, position)

                # Resize the modified image for embedding in the email
                resized_img = modified_img.resize((326, 503))

                # Save the combined image
                resized_img.save(output_path)

        print("Barcode added to design successfully and saved as PNG.")

        # Email content
        expiration_date = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime("%m/%d/%Y")
        email_subject = "Fietsstation Voucher Code"
        email_body = f"""<html>
                            <body>
                                <p>Beste {name},</p></p>
                                <p>Enorm bedankt van het doneren van uw fiets!</p>
                                <p>Hiervoor heeft u een voucher code voor 10% Korting! op een nieuwe fiets verdient.</p> 
                                <p>Dit is uw Coupon code <b>{code_without_asterisk}</b><p>
                                <p>Deze is geldig tot {expiration_date}</p>
                                <p>Je kunt ook de foto onderaan gebruiken als je hem komt inleveren.</p>
                                <p>ICT Medewerker Beginstation</p>
                                <img src="cid:image1">
                            </body>
                        </html>"""

        # Create Outlook application object
        outlook = win32.Dispatch('outlook.application')
        
        # Create a new mail item
        mail = outlook.CreateItem(0)
        mail.Subject = email_subject
        mail.HTMLBody = email_body

        # Add recipient email address
        mail.To = email_address

        # Set image as inline content
        attachment = mail.Attachments.Add(output_path)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "image1")

        # Send the email
        mail.Send()

        print("Email sent with the voucher code and image embedded in the body.")
        os.remove(barcode_filepath)  # Delete the barcode PNG file

    except Exception as e:
        print(f"An error occurred: {str(e)}")

def get_random_unactivated_code():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1

        # Retrieve all data in one request
        data = sheet.get_all_values()

        unactivated_codes = [row[0] for row in data[1:] if row[0] and row[0].startswith("*") and not row[2]]
        if unactivated_codes:
            return random.choice(unactivated_codes)
        else:
            return None
    except SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
        return None
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

def delete_all_info():
    password = "Bstation7173"
    entered_password = input("Enter password to confirm: ")
    
    if entered_password == password:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)

        try:
            sheet = client.open('voucher-barcode1').sheet1
            # Fetch all values from the sheet
            values = sheet.get_all_values()

            if len(values) >= 2:
                # Delete all rows except the header (keeping first row intact)
                sheet.delete_rows(2, len(values) - 0)
                print("All information (except headers) deleted successfully.")
            else:
                print("No information to delete.")
        except SpreadsheetNotFound:
            print("Spreadsheet 'voucher-barcode1' not found.")
        except Exception as e:
            print(f"An error occurred: {str(e)}")
    else:
        print("Incorrect password. Aborting operation.")

def main():
    while True:
        clear_screen()
        print("\nMenu:")
        print("1. Generate codes")
        print("2. Check code validity")
        print("3. Print")
        print("4. Info status")
        print("5. Activate code")
        print("6. Redeem code")
        print("7. Save as PNG")
        print("8. Activate and email")
        print("9. Exit")

        choice = input("Enter your choice: ")

        if choice == '1':
            num_codes = int(input("How many codes do you want to generate? "))
            clear_screen()
            add_to_google_sheet(num_codes)
        elif choice == '2':
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
            code = input("Enter code to activate: ")
            clear_screen()
            activate_code(code)
        elif choice == '6':
            code = input("Enter code to redeem: ")
            clear_screen()
            redeem_code(code)
        elif choice == '7':
            code = input("Enter code to save as PNG: ")
            clear_screen()
            save_as_png(code)
        elif choice == '8':
            name = input("Enter your name: ")
            email_address = input("Enter your email address: ")
            clear_screen()
            activate_and_email(name, email_address)
        elif choice == '9':
            print("Exiting program...")
            break
        elif choice == '7173':
            clear_screen()
            delete_all_info()  # Secret function for deleting all information
        else:
            print("Invalid choice. Please try again.")
        
        time.sleep(3)

if __name__ == "__main__":
    main()
