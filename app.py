import gspread
import time
import os
import subprocess
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from gspread.exceptions import SpreadsheetNotFound
from barcode import Code128
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import base64
import random
import string
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.clock import Clock
from barcode import generate
from barcode.writer import ImageWriter
import img2pdf
import re

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
    code128 = barcode.get_barcode_class('code128')
    ean = code128(code, writer=ImageWriter())
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

            # Specify the scopes for the credentials
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

            # Pass the scopes to the from_json_keyfile_name method
            creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file_path, scopes)
            get_credentials._creds = creds
        except Exception as e:
            print(f"Error bij het laden van credentials: {str(e)}")
            show_success_popup("Warning","Error with loading credentials")
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

        print("Connected to Google Sheets successfully.")

        start_index = sheet.row_count + 1

        print(f"Start index: {start_index}")

        # Generate and add new codes
        batch_data = []
        for _ in range(num_codes):
            unique_code = '*' + generate_unique_code()  # Add * to indicate inactive code
            batch_data.append([unique_code])

        print("Batch data generated successfully.")

        # Add the batch data to the spreadsheet
        sheet.append_rows(batch_data)

        print(f"{num_codes} codes generated and added to the spreadsheet.")
        show_success_popup("Success", f"{num_codes} codes generated and added to the spreadsheet.")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' not found.")
        show_success_popup("Warning","Cant find voucher-barcode1 spreadsheet!")
    except Exception as e:
        print(f"An error occurred: {str(e)}")  # Print the actual exception message

def show_success_popup(title, message):
    layout = BoxLayout(orientation='vertical')
    label = Label(text=message)
    close_button = Button(text='Close', on_press=lambda x: success_popup.dismiss())
    layout.add_widget(label)
    layout.add_widget(close_button)
    success_popup = Popup(title=title, content=layout, size_hint=(0.75, 0.5))
    success_popup.open()

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
                show_success_popup("Info","Code was already active/redeemed")
            elif sheet.cell(cell.row, 3).value:  # Check if the code has been marked as printed
                sheet.update_cell(cell.row, 1, cell.value[1:])  # Remove the * from the code
                expiration_date = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime("%m/%d/%Y")
                sheet.update_cell(cell.row, 4, expiration_date)  # Update expiration date in column D
                print(f"Code geactiveerd. is geldig tot {expiration_date}.")
                show_success_popup("Success", f"Code geactiveerd. is geldig tot {expiration_date}.")
            else:
                print("Code is niet geprint. Activatie is niet mogelijk.")
                show_success_popup("Warning","Cant activate a unprinted code!")
        else:
            print("Code niet gevonden.")
            show_success_popup("Error","Code has not been found")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' niet gevonden.")
        show_success_popup("Error","Spreadsheet niet gevonden")
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
                    print("Code is ingeleverd")
                    show_success_popup("Info", "Code is ingeleverd")
                    time.sleep(3)
                elif code_value.startswith("*"):
                    print("Code is niet geactiveerd.")
                    show_success_popup("Info","Code is unactive")
                else:
                    expiration_date = row[3]
                    print(f"Code is geldig en geactiveerd. is geldig tot {expiration_date}.")
                    show_success_popup("Info", f"Code is geldig en geactiveerd. is geldig tot {expiration_date}.")
                    return expiration_date  # Return the expiration date
                break
        else:
            print("Code niet gevonden.")
            show_success_popup("Error", "Code niet gevonden.")
            return None  # Return None when code is not found
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' niet gevonden.")
    except Exception as e:
        print(f"een error is gebeurd: {str(e)}")
        show_success_popup("Error", f"Een error is gebeurd: {str(e)}")

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
            font_path = os.path.join(DEFAULT_DIR, 'arial.ttf')
            font = ImageFont.truetype(font_path, 35)
            draw = ImageDraw.Draw(stretched_img)
            text_bbox = draw.textbbox((0, 0), code_without_asterisk, font=font)
            text_x = position[0] + (barcode_img.width - text_bbox[2]) / 2
            text_y = position[1] + barcode_img.height + 5
            draw.text((text_x, text_y), code_without_asterisk, font=font, fill=(0, 0, 0))

            output_filename = f"coupon_{code_without_asterisk}.png"
            output_path = os.path.join(OUTPUT_DIR, output_filename)
            stretched_img.save(output_path)
            print(f"Coupon opgeslagen als: {output_path}")
            print_image_in_portrait(output_path)

            return output_path

    except Exception as e:
        print(f"Een error is gebeurd: {e}")
        return None

def make_printed(credentials_file_path=None):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Failed to load credentials. Exiting.")
            return

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        values = sheet.get_all_values()

        # Fetch unprinted codes
        unprinted_codes = [(row[0], i + 1) for i, row in enumerate(values[1:]) if len(row) >= 3 and row[0].startswith("*") and row[2] != "Printed"]

        if len(unprinted_codes) < 12:
            print("Niet genoeg onactieve codes beschikbaar.")
            show_success_popup("Info","Please generate at least 12 more codes")
            time.sleep(1)
            show_success_popup("Info","Not enough codes to use")
            return

        selected_codes = random.sample(unprinted_codes, 12)
        for code, row_index in selected_codes:
            sheet.update_cell(row_index + 1, 3, "Printed")
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

            coupon_width, coupon_height = 319, 509
            positions = [(x * coupon_width, y * coupon_height) for y in range(3) for x in range(4)]

            for code, row_index in selected_codes:
                image_filename = os.path.join(DEFAULT_DIR, 'standaard.png')
                if not os.path.exists(image_filename):
                    print("Error: Foto bestand niet gevonden.")
                    return

                with Image.open(image_filename) as img:
                    code_without_asterisk = code.replace('*', '')
                    barcode = Code128(code_without_asterisk, writer=ImageWriter())
                    barcode_img = barcode.render()

                    barcode_width, barcode_height = 340, 75
                    barcode_x, barcode_y = -10, 350

                    barcode_img = barcode_img.resize((barcode_width, barcode_height))
                    img.paste(barcode_img, (barcode_x, barcode_y))

                    coupon_img = img.resize((coupon_width, coupon_height))

                    pos_x, pos_y = positions[selected_codes.index((code, row_index))]
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
            
            # Code for Android to open the PDF
            try:
                import androidhelper
                droid = androidhelper.Android()
                droid.startActivity('android.intent.action.VIEW', 'application/pdf', coupon_list_pdf)
                print(f"PDF geopend: {coupon_list_pdf}")
            except ImportError:
                print(f"PDF is opgeslagen op: {coupon_list_pdf}. Open het handmatig.")
                show_success_popup("Pdf has been made","A pdf has been put in the output file")
    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")

# Ensure the directories exist
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
if not os.path.exists(DEFAULT_DIR):
    os.makedirs(DEFAULT_DIR)

def mark_code_as_redeemed(code, credentials_file_path=CREDENTIALS_FILE_PATH):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Gefaald om credentials te laden. Sluiten.")
            return

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode1').sheet1
        cell = sheet.find(code)

        if cell:
            if cell.value.startswith("~"):
                print("Code is al ingeleverd.")
                show_success_popup("Info","Code is al ingeleverd")
            elif cell.value.startswith("*"):
                print("Code is unactive.")
                show_success_popup("Warning","Cant redeem a unactive code!")
            else:
                sheet.update_cell(cell.row, 1, "~" + cell.value)
                redemption_date = datetime.datetime.now().strftime("%m/%d/%Y")
                sheet.update_cell(cell.row, 4, redemption_date)  # Update redemption date in column D
                print(f"Code {code} gemarkeerd als ingeleverd op {redemption_date}.")
                show_success_popup("Success",f"Code {code} gemarkeerd als ingeleverd op {redemption_date}.")
        else:
            print("Code niet gevonden.")
            show_success_popup("Error","Code niet gevonden")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode1' niet gevonden.")
        show_success_popup("Error","Spreadsheet niet gevonden")
    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")

class VoucherApp(App):
    def build(self):
        self.screen_manager = ScreenManager()

        self.main_screen = MainScreen(name='main')
        self.generate_screen = GenerateScreen(name='generate')
        self.activate_screen = ActivateScreen(name='activate')
        self.redeem_screen = RedeemScreen(name='redeem')
        self.print_screen = PrintScreen(name='print')
        self.check_screen = CheckScreen(name='check')

        self.screen_manager.add_widget(self.main_screen)
        self.screen_manager.add_widget(self.generate_screen)
        self.screen_manager.add_widget(self.activate_screen)
        self.screen_manager.add_widget(self.redeem_screen)
        self.screen_manager.add_widget(self.print_screen)
        self.screen_manager.add_widget(self.check_screen)

        return self.screen_manager

class MainScreen(Screen):
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical')

        self.generate_button = Button(text='Generate Vouchers', on_press=self.switch_to_generate)
        self.activate_button = Button(text='Activate Voucher', on_press=self.switch_to_activate)
        self.redeem_button = Button(text='Redeem Voucher', on_press=self.switch_to_redeem)
        self.print_button = Button(text='Print Vouchers', on_press=self.switch_to_print)
        self.check_button = Button(text='Check Voucher', on_press=self.switch_to_check)

        layout.add_widget(self.generate_button)
        layout.add_widget(self.activate_button)
        layout.add_widget(self.redeem_button)
        layout.add_widget(self.print_button)
        layout.add_widget(self.check_button)

        self.add_widget(layout)

    def switch_to_generate(self, instance):
        self.manager.current = 'generate'

    def switch_to_activate(self, instance):
        self.manager.current = 'activate'

    def switch_to_redeem(self, instance):
        self.manager.current = 'redeem'

    def switch_to_print(self, instance):
        self.manager.current = 'print'

    def switch_to_check(self, instance):
        self.manager.current = 'check'

class GenerateScreen(Screen):
    def __init__(self, **kwargs):
        super(GenerateScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical')

        self.generate_input = TextInput(hint_text='Enter number of vouchers to generate', multiline=False)
        self.generate_button = Button(text='Generate Vouchers', on_press=self.generate_vouchers)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main)

        layout.add_widget(self.generate_input)
        layout.add_widget(self.generate_button)
        layout.add_widget(self.back_button)

        self.add_widget(layout)

    def generate_vouchers(self, instance):
        instance.background_color = [0, 0, 1, 1]  # Change button to blue
        Clock.schedule_once(lambda dt: self._generate_vouchers(instance), 0.1)

    def _generate_vouchers(self, instance):
        try:
            num_codes = int(self.generate_input.text)
            add_to_google_sheet(num_codes)
        except ValueError:
            self.show_success_popup("Error", "Invalid number entered.")
        finally:
            instance.background_color = [1, 1, 1, 1]  # Revert button color

    def switch_to_main(self, instance):
        self.manager.current = 'main'

    def show_success_popup(self, title, message):
        layout = BoxLayout(orientation='vertical')
        label = Label(text=message)
        close_button = Button(text='Close', on_press=lambda x: popup.dismiss())
        layout.add_widget(label)
        layout.add_widget(close_button)
        popup = Popup(title=title, content=layout, size_hint=(0.75, 0.5))
        popup.open()

class ActivateScreen(Screen):
    def __init__(self, **kwargs):
        super(ActivateScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical')

        self.activate_input = TextInput(hint_text='Enter voucher code to activate', multiline=False)
        self.activate_button = Button(text='Activate Voucher', on_press=self.activate_voucher)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main)

        layout.add_widget(self.activate_input)
        layout.add_widget(self.activate_button)
        layout.add_widget(self.back_button)

        self.add_widget(layout)

    def activate_voucher(self, instance):
        instance.background_color = [0, 0, 1, 1]  # Change button to blue
        Clock.schedule_once(lambda dt: self._activate_voucher(instance), 0.1)

    def _activate_voucher(self, instance):
        code = self.activate_input.text.strip()
        if code:
            activate_code(code)
        else:
            self.show_success_popup("Error", "Enter a voucher code to activate.")
        instance.background_color = [1, 1, 1, 1]  # Revert button color

    def switch_to_main(self, instance):
        self.manager.current = 'main'

    def show_success_popup(self, title, message):
        layout = BoxLayout(orientation='vertical')
        label = Label(text=message)
        close_button = Button(text='Close', on_press=lambda x: popup.dismiss())
        layout.add_widget(label)
        layout.add_widget(close_button)
        popup = Popup(title=title, content=layout, size_hint=(0.75, 0.5))
        popup.open()

class RedeemScreen(Screen):
    def __init__(self, **kwargs):
        super(RedeemScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical')

        self.redeem_input = TextInput(hint_text='Enter voucher code to redeem', multiline=False)
        self.redeem_button = Button(text='Redeem Voucher', on_press=self.redeem_voucher)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main)

        layout.add_widget(self.redeem_input)
        layout.add_widget(self.redeem_button)
        layout.add_widget(self.back_button)

        self.add_widget(layout)

    def redeem_voucher(self, instance):
        instance.background_color = [0, 0, 1, 1]  # Change button to blue
        Clock.schedule_once(lambda dt: self._redeem_voucher(instance), 0.1)

    def _redeem_voucher(self, instance):
        code = self.redeem_input.text.strip()
        if code:
            mark_code_as_redeemed(code)
        else:
            self.show_success_popup("Error", "Enter a voucher code to redeem.")
        instance.background_color = [1, 1, 1, 1]  # Revert button color

    def switch_to_main(self, instance):
        self.manager.current = 'main'

    def show_success_popup(self, title, message):
        layout = BoxLayout(orientation='vertical')
        label = Label(text=message)
        close_button = Button(text='Close', on_press=lambda x: popup.dismiss())
        layout.add_widget(label)
        layout.add_widget(close_button)
        popup = Popup(title=title, content=layout, size_hint=(0.75, 0.5))
        popup.open()

class PrintScreen(Screen):
    def __init__(self, **kwargs):
        super(PrintScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical')

        self.print_button = Button(text='Print Vouchers', on_press=self.print_vouchers)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main)

        layout.add_widget(self.print_button)
        layout.add_widget(self.back_button)

        self.add_widget(layout)

    def print_vouchers(self, instance):
        instance.background_color = [0, 0, 1, 1]  # Change button to blue
        Clock.schedule_once(lambda dt: self._print_vouchers(instance), 0.1)

    def _print_vouchers(self, instance):
        make_printed()
        instance.background_color = [1, 1, 1, 1]  # Revert button color

    def switch_to_main(self, instance):
        self.manager.current = 'main'

class CheckScreen(Screen):
    def __init__(self, **kwargs):
        super(CheckScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical')

        self.check_input = TextInput(hint_text='Enter voucher code to check', multiline=False)
        self.check_button = Button(text='Check Voucher', on_press=self.check_voucher)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main)

        layout.add_widget(self.check_input)
        layout.add_widget(self.check_button)
        layout.add_widget(self.back_button)

        self.add_widget(layout)

    def check_voucher(self, instance):
        instance.background_color = [0, 0, 1, 1]  # Change button to blue
        Clock.schedule_once(lambda dt: self._check_voucher(instance), 0.1)

    def _check_voucher(self, instance):
        code = self.check_input.text.strip()
        if code:
            check_code_validity(code)
        else:
            self.show_success_popup("Error", "Enter a voucher code to check.")
        instance.background_color = [1, 1, 1, 1]  # Revert button color

    def switch_to_main(self, instance):
        self.manager.current = 'main'

    def show_success_popup(self, title, message):
        layout = BoxLayout(orientation='vertical')
        label = Label(text=message)
        close_button = Button(text='Close', on_press=lambda x: popup.dismiss())
        layout.add_widget(label)
        layout.add_widget(close_button)
        popup = Popup(title=title, content=layout, size_hint=(0.75, 0.5))
        popup.open()

if __name__ == '__main__':
    VoucherApp().run()
