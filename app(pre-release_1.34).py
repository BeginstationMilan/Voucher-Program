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
from kivy.utils import get_color_from_hex 
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
import random
import string
from PIL import Image
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
from kivy.uix.gridlayout import GridLayout
from kivy.uix.spinner import Spinner
from kivy.uix.image import Image as KivyImage
import img2pdf
import re
from google.oauth2.credentials import Credentials
from kivy.uix.screenmanager import ScreenManager, SlideTransition
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.core.window import Window
from kivy.uix.popup import Popup
from kivy.uix.dropdown import DropDown
from dotenv import load_dotenv
import webbrowser
import urllib.parse
from kivy.uix.image import Image as KivyImage
from kivy.uix.image import Image
from kivy.uix.relativelayout import RelativeLayout
import sys
from PIL import Image, ImageDraw, ImageFont
from PIL import Image as PILImage
from kivy.uix.image import Image


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
        sheet = client.open('voucher-barcode').sheet1

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
        print("Spreadsheet 'voucher-barcode' not found.")
        show_success_popup("Warning","Cant find voucher-barcode spreadsheet!")
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
        sheet = client.open('voucher-barcode').sheet1
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
        print("Spreadsheet 'voucher-barcode' niet gevonden.")
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
        sheet = client.open('voucher-barcode').sheet1

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
        print("Spreadsheet 'voucher-barcode' niet gevonden.")
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

        script_dir = os.path.dirname(os.path.abspath(__file__))
        DEFAULT_DIR = os.path.join(script_dir, 'Default')
        OUTPUT_DIR = os.path.join(script_dir, 'Output')

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode').sheet1
        values = sheet.get_all_values()

        unprinted_codes = [(row[0], i + 1) for i, row in enumerate(values[1:]) if len(row) >= 3 and row[0].startswith("*") and row[2] != "Printed"]

        if len(unprinted_codes) < 12:
            print("Not enough inactive codes available.")
            show_error_popup("Error", "Not enough inactive codes available. Please generate more codes.")
            return

        selected_codes = random.sample(unprinted_codes, 12)
        for code, row_index in selected_codes:
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

            coupon_width, coupon_height = 319, 509
            positions = [(x * coupon_width, y * coupon_height) for y in range(3) for x in range(4)]

            for code, row_index in selected_codes:
                image_filename = os.path.join(DEFAULT_DIR, 'standaard.png')
                if not os.path.exists(image_filename):
                    print("Error: Image file not found.")
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
            print("\nPage generation complete.")

            images = [final_output_path, os.path.join(DEFAULT_DIR, 'standaard_full.png')]
            pdf_bytes = img2pdf.convert(images)

            with open(coupon_list_pdf, "wb") as f:
                f.write(pdf_bytes)

            print("Coupon List PDF generated.")

            app = App.get_running_app()
            app.screen_manager.current = 'print'
            print("Navigated to the print screen.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        show_error_popup("Error", str(e))

class MainScreen(Screen):
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical')
        self.label = Label(text='Press the button to generate coupon list')
        self.button = Button(text='Generate', on_press=self.generate_coupon_list)
        layout.add_widget(self.label)
        layout.add_widget(self.button)
        self.add_widget(layout)

    def generate_coupon_list(self, instance):
        self.label.text = "Generating..."
        make_printed()
        self.label.text = "Coupon list generated. Check the Output directory."

class PrintScreen(Screen):
    def __init__(self, **kwargs):
        super(PrintScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical')
        self.label = Label(text='Coupons printed successfully.')
        self.add_widget(self.label)

class MyApp(App):
    def build(self):
        self.screen_manager = ScreenManager()
        self.main_screen = MainScreen(name='main')
        self.print_screen = PrintScreen(name='print')

        self.screen_manager.add_widget(self.main_screen)
        self.screen_manager.add_widget(self.print_screen)

        return self.screen_manager

    def show_error_popup(self, title, message):
        box = BoxLayout(orientation='vertical')
        box.add_widget(Label(text=message))
        btn = Button(text='Close', size_hint_y=0.2)
        box.add_widget(btn)
        popup = Popup(title=title, content=box, size_hint=(0.8, 0.5))
        btn.bind(on_press=popup.dismiss)
        popup.open()

if __name__ == '__main__':
    # Ensure the directories exist
    script_dir = os.path.dirname(os.path.abspath(__file__))
    DEFAULT_DIR = os.path.join(script_dir, 'Default')
    OUTPUT_DIR = os.path.join(script_dir, 'Output')

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
    if not os.path.exists(DEFAULT_DIR):
        os.makedirs(DEFAULT_DIR)

def show_error_popup(title, message):
    layout = BoxLayout(orientation='vertical')
    label = Label(text=message)
    close_button = Button(text='Close', on_press=lambda x: popup.dismiss())
    layout.add_widget(label)
    layout.add_widget(close_button)
    popup = Popup(title=title, content=layout, size_hint=(0.75, 0.5))
    popup.open()

# Ensure the directories exist
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
if not os.path.exists(DEFAULT_DIR):
    os.makedirs(DEFAULT_DIR)

def open_file_with_path(file_path):
    try:
        os.system(f"open {file_path}")
        print(f"Opened file: {file_path}")
    except Exception as e:
        print(f"Failed to open file: {str(e)}")
        show_error_popup("Error", "Failed to open file.")

def mark_code_as_redeemed(code, credentials_file_path=CREDENTIALS_FILE_PATH):
    try:
        creds = get_credentials(credentials_file_path)
        if creds is None:
            print("Gefaald om credentials te laden. Sluiten.")
            return

        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode').sheet1
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
                sheet.update_cell(cell.row, 5, redemption_date)  # Update redemption date in column E
                print(f"Code {code} gemarkeerd als ingeleverd op {redemption_date}.")
                show_success_popup("Success",f"Code {code} gemarkeerd als ingeleverd op {redemption_date}.")
        else:
            print("Code niet gevonden.")
            show_success_popup("Error","Code niet gevonden")
    except gspread.exceptions.SpreadsheetNotFound:
        print("Spreadsheet 'voucher-barcode' niet gevonden.")
        show_success_popup("Error","Spreadsheet niet gevonden")
    except Exception as e:
        print(f"Een error is gebeurd: {str(e)}")

from functools import wraps
from threading import Thread

def run_in_thread(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        t = Thread(target=func, args=args, kwargs=kwargs)
        t.start()
    return wrapper

class LoginScreen(Screen):
    def __init__(self, **kwargs):
        super(LoginScreen, self).__init__(**kwargs)
        Window.clearcolor = get_color_from_hex('#3A7ED2')

        layout = BoxLayout(orientation='vertical', spacing=10, padding=20)

        # Set the path for login image
        login_image_path = os.path.join('GUI', 'Login_image.png')
        try:
            # Create the login image
            login_image = Image(source=login_image_path)
            layout.add_widget(login_image)
        except Exception as e:
            print(f"Error loading login image: {e}")

        self.username_input = TextInput(hint_text='Enter username', multiline=False, font_size='23sp', size_hint_y=None, height=60)
        self.password_input = TextInput(hint_text='Enter password', multiline=False, password=True, font_size='23sp', size_hint_y=None, height=60)

        layout.add_widget(self.username_input)
        layout.add_widget(self.password_input)

        # Set the paths for login button
        login_button_path_normal = os.path.join('GUI', 'Login_Button.png')
        login_button_path_down = os.path.join('GUI', 'Login_Button1.png')

        button_layout = RelativeLayout(size_hint=(1, None), height=160)

        try:
            # Create the login button
            self.login_button = Button(background_normal=login_button_path_normal, background_down=login_button_path_down, size_hint=(1, None), height=170, on_press=self.login)

            login_label = Label(text="Login", font_size='37.5sp', bold=True, color=(1, 1, 1, 1), pos_hint={'center_x': 0.5, 'center_y': 0.5})

            button_layout.add_widget(self.login_button)
            button_layout.add_widget(login_label)

            layout.add_widget(button_layout)
        except Exception as e:
            print(f"Error loading login button: {e}")

        self.add_widget(layout)

    def on_pre_enter(self, *args):
        self.username_input.text = ''
        self.password_input.text = ''

    def login(self, instance):
        username = self.username_input.text.strip()
        password = self.password_input.text.strip()
        self.authenticate(username, password)

    def after_authentication(self, username, account_type):
        if account_type:
            self.manager.get_screen('main').set_logged_in_type(account_type)
            self.switch_to_main()
        else:
            self.show_error_popup("Error", "Invalid credentials")

    def authenticate(self, username, password):
        Clock.schedule_once(lambda dt: self._authenticate(username, password), 0)

    def _authenticate(self, username, password):
        try:
            client = gspread.authorize(get_credentials())
            sheet = client.open('accounts').sheet1
        except gspread.exceptions.SpreadsheetNotFound:
            print("Spreadsheet 'accounts' not found.")
            self.after_authentication(username, None)
            return

        user_data = sheet.get_all_records()

        for user in user_data:
            if user['username'] == username and user['password'] == password:
                if 'type' in user:
                    self.after_authentication(username, user['type'])
                    return
                else:
                    print("Account type not found for user:", username)
                    self.after_authentication(username, None)
                    return
        self.after_authentication(username, None)

    def set_logged_in_type(self, account_type):
        self.logged_in_type = account_type

    def switch_to_main(self):
        self.manager.current = 'main'

    def show_error_popup(self, title, message):
        popup = Popup(title=title, content=Label(text=message), size_hint=(None, None), size=(400, 200))
        popup.open()


class VoucherApp(App):
    def build(self):
        self.screen_manager = ScreenManager(transition=SlideTransition())

        self.login_screen = LoginScreen(name='login')
        self.main_screen = MainScreen(name='main')
        self.generate_screen = GenerateScreen(name='generate')
        self.activate_screen = ActivateScreen(name='activate')
        self.redeem_screen = RedeemScreen(name='redeem')
        self.print_screen = PrintScreen(name='print')
        self.check_screen = CheckScreen(name='check')  # Add CheckScreen here
        self.code_control_screen = CodeControlScreen(name='code_control')
        self.email_screen = EmailScreen(name='email')

        self.screen_manager.add_widget(self.login_screen)
        self.screen_manager.add_widget(self.main_screen)
        self.screen_manager.add_widget(self.generate_screen)
        self.screen_manager.add_widget(self.activate_screen)
        self.screen_manager.add_widget(self.redeem_screen)
        self.screen_manager.add_widget(self.print_screen)
        self.screen_manager.add_widget(self.check_screen)  # Add CheckScreen here
        self.screen_manager.add_widget(self.code_control_screen)
        self.screen_manager.add_widget(self.email_screen)

        self.screen_manager.current = 'login'

        Window.clearcolor = get_color_from_hex('#3A7ED2')

        return self.screen_manager


class MainScreen(Screen):
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        Window.clearcolor = get_color_from_hex('#00BFFF')  # Bright blue background color

        self.layout = GridLayout(cols=2, spacing=10, padding=20)
        button_color = get_color_from_hex('#3D7CC9')  # Consistent color with EmailScreen

        self.buttons = {
            'Generate': Button(text='Generate Vouchers', on_press=self.switch_to_generate,
                               background_color=button_color, font_size='18sp'),
            'Activate': Button(text='Activate Voucher', on_press=self.switch_to_activate,
                               background_color=button_color, font_size='18sp'),
            'Redeem': Button(text='Redeem Voucher', on_press=self.switch_to_redeem,
                             background_color=button_color, font_size='18sp'),
            'Print': Button(text='Print Vouchers', on_press=self.switch_to_print,
                            background_color=button_color, font_size='18sp'),
            'Check': Button(text='Check Voucher', on_press=self.switch_to_check,
                            background_color=button_color, font_size='18sp'),
            'CodeControl': Button(text='Code Control', on_press=self.switch_to_code_control,
                                  background_color=button_color, font_size='18sp'),
            'Email': Button(text='Email', on_press=self.switch_to_email,
                            background_color=button_color, font_size='18sp'),
            'Return': Button(text='Return to Login', on_press=self.switch_to_login,
                             background_color=button_color, font_size='18sp')
        }

        for button in self.buttons.values():
            self.layout.add_widget(button)

        self.add_widget(self.layout)

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

    def switch_to_code_control(self, instance):
        self.manager.current = 'code_control'

    def switch_to_email(self, instance):
        self.manager.current = 'email'

    def switch_to_login(self, instance):
        self.manager.current = 'login'

    def set_logged_in_type(self, user_type):
        # Reset all buttons to enabled state
        for button in self.buttons.values():
            button.disabled = False

        # Disable buttons based on user_type
        if user_type == 'Work':
            self.buttons['Generate'].disabled = True
            self.buttons['Print'].disabled = True
            self.buttons['CodeControl'].disabled = True
        elif user_type == 'Customer':
            self.buttons['Generate'].disabled = True
            self.buttons['Activate'].disabled = True
            self.buttons['Print'].disabled = True
            self.buttons['CodeControl'].disabled = True
            self.buttons['Email'].disabled = True  # Disable Email button for customer type


class GenerateScreen(Screen):
    def __init__(self, **kwargs):
        super(GenerateScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', spacing=10, padding=20)

        self.generate_input = TextInput(hint_text='Enter number of vouchers to generate', multiline=False, font_size=30, hint_text_color=[0.8, 0.8, 0.8, 1])
        self.generate_button = Button(text='Generate Vouchers', on_press=self.generate_vouchers, background_color=[0.3, 0.5, 0.7, 1], font_size=30)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main, background_color=[0.7, 0.3, 0.3, 1], font_size=25)

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


class EmailScreen(Screen):
    def __init__(self, **kwargs):
        super(EmailScreen, self).__init__(**kwargs)
        Window.clearcolor = get_color_from_hex('#101820')  # Dark background color

        self.layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        self.label = Label(text="Enter email address:", color=get_color_from_hex('#F2AA4C'), font_size='20sp')
        self.layout.add_widget(self.label)

        self.email_input = TextInput(hint_text="Email", multiline=False, padding_y=(5, 5), size_hint_y=None, height=120)
        self.layout.add_widget(self.email_input)

        self.activate_button = Button(text="Activate Voucher", size_hint=(1, 0.3), background_color=get_color_from_hex('#3D7CC9'), color=get_color_from_hex('#FFFFFF'))
        self.activate_button.bind(on_release=self.activate_voucher)
        self.layout.add_widget(self.activate_button)

        self.return_button = Button(text="Return to Main Menu", size_hint=(1, 0.3), background_color=get_color_from_hex('#3D7CC9'), color=get_color_from_hex('#FFFFFF'))
        self.return_button.bind(on_release=self.return_to_menu)
        self.layout.add_widget(self.return_button)  # Fixed this line

        self.add_widget(self.layout)

    def activate_voucher(self, instance):
        self.activate_button.text = "Activating..."
        Clock.schedule_once(self._activate_voucher, 0.1)

    def _activate_voucher(self, dt):
        codes = self.fetch_codes()
        if not codes:
            self.show_error_popup("Not Enough codes usable")
            self.activate_button.text = "Activate Voucher"
            return

        selected_code = codes[0]
        self.update_sheet(selected_code)
        self.email_code(selected_code)
        self.activate_button.text = "Activate Voucher"

    def fetch_codes(self):
        creds = self.get_credentials()
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode').sheet1
        codes = sheet.get_all_values()[1:]  # Skip header row
        valid_codes = [code[0] for code in codes if code[0].startswith('*') and not code[2]]
        return valid_codes  # Corrected from 'return valid' to 'return valid_codes'

    def update_sheet(self, code):
        creds = self.get_credentials()
        client = gspread.authorize(creds)
        sheet = client.open('voucher-barcode').sheet1
        cell = sheet.find(code)
        expiration_date = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime("%m/%d/%Y")
        sheet.update_cell(cell.row, 1, code[1:])  # Remove '*'
        sheet.update_cell(cell.row, 3, "Emailed")
        sheet.update_cell(cell.row, 4, expiration_date)  # Add expiration date to Column D

    def email_code(self, code):
        code_without_asterisk = code[1:] if code.startswith("*") else code
        expiration_date = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime("%d-%m-%Y")
        email_body = f"""Beste klant,

Bedankt voor het doneren van uw fiets. Voor dit krijgt u een coupon code van 10% korting.

Vouchercode: {code_without_asterisk}
Geldig tot: {expiration_date}

Bewaar deze code veilig u mag deze ook weggeven aan een ander persoon.

Met vriendelijke groeten,
Het Fietsstation"""
        encoded_email_body = urllib.parse.quote(email_body)
        webbrowser.open(f"mailto:{self.email_input.text}?subject=Fietsstation Vouchercode&body={encoded_email_body}")

    def show_error_popup(self, message):
        popup = Popup(title='Error', content=Label(text=message), size_hint=(None, None), size=(400, 400))
        popup.open()

    def return_to_menu(self, instance):
        app_instance = App.get_running_app()
        screen_manager = app_instance.root
        if screen_manager:
            screen_manager.current = "main"

    def get_credentials(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        credentials_file_path = os.path.join(script_dir, 'credentials.json')
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        return ServiceAccountCredentials.from_json_keyfile_name(credentials_file_path, scopes)

class RedeemScreen(Screen):
    def __init__(self, **kwargs):
        super(RedeemScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', spacing=10, padding=20)

        self.redeem_input = TextInput(hint_text='Enter voucher code to redeem', multiline=False, font_size=30)
        self.redeem_button = Button(text='Redeem Voucher', on_press=self.redeem_voucher, background_color=[0.3, 0.5, 0.7, 1], font_size=30)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main, background_color=[0.7, 0.3, 0.3, 1], font_size=30)

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
        layout = BoxLayout(orientation='vertical', spacing=10, padding=20)

        self.print_button = Button(text='Print Vouchers', on_press=self.print_vouchers, background_color=[0.3, 0.5, 0.7, 1], font_size=30)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main, background_color=[0.7, 0.3, 0.3, 1], font_size=30)

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
        layout = BoxLayout(orientation='vertical', spacing=10, padding=20)

        self.check_input = TextInput(hint_text='Enter voucher code to check', multiline=False, font_size=30)
        self.check_button = Button(text='Check Voucher', on_press=self.check_voucher, background_color=[0.3, 0.5, 0.7, 1], font_size=30)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main, background_color=[0.7, 0.3, 0.3, 1], font_size=30)

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

    def show_success_popup(title, message):
        layout = BoxLayout(orientation='vertical', spacing=10, padding=20)
        label = Label(text=message, font_size=14)
        close_button = Button(text='Close', on_press=lambda x: success_popup.dismiss(), background_color=[0.6, 0.3, 0.3, 1], font_size=14)
        layout.add_widget(label)
        layout.add_widget(close_button)
        success_popup = Popup(title=title, content=layout, size_hint=(0.75, 0.5))
        success_popup.open()


from kivy.clock import Clock

class ActivateScreen(Screen):
    def __init__(self, **kwargs):
        super(ActivateScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', spacing=10, padding=20)

        self.activate_input = TextInput(hint_text='Enter voucher code to activate', multiline=False, font_size=30)
        self.activate_button = Button(text='Activate Voucher', on_press=self.activate_voucher, background_color=[0.3, 0.5, 0.7, 1], font_size=30)
        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main, background_color=[0.7, 0.3, 0.3, 1], font_size=30)

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

import gspread
from kivy.app import App
from kivy.uix.screenmanager import Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView
from kivy.clock import Clock
from oauth2client.service_account import ServiceAccountCredentials
from kivy.uix.popup import Popup
from kivy.uix.label import Label

class CodeControlScreen(Screen):
    def __init__(self, **kwargs):
        super(CodeControlScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', spacing=10, padding=20)

        self.refresh_button = Button(text='Refresh', on_press=self._refresh_codes,
                                     background_color=[0.3, 0.7, 0.3, 1], font_size=14, size_hint=(1, None), height=80)

        self.search_input = TextInput(hint_text='Search code', multiline=False, font_size=23, size_hint=(1, None), height=80)
        self.search_input.bind(on_text_validate=self.update_code_buttons)

        self.status_filter = Spinner(
            text='All',
            values=('All', 'Active', 'Unactive', 'Redeemed'),
            size_hint=(1, None),
            height=50,
            sync_height=True
        )
        self.status_filter.bind(text=self.update_code_buttons)

        self.code_list_layout = BoxLayout(orientation='vertical', spacing=10, padding=20, size_hint_y=None)
        self.code_list_layout.bind(minimum_height=self.code_list_layout.setter('height'))

        scroll_view = ScrollView(size_hint=(1, 1))
        scroll_view.add_widget(self.code_list_layout)

        self.back_button = Button(text='Back to Main Menu', on_press=self.switch_to_main,
                                  background_color=[0.7, 0.3, 0.3, 1], font_size=14, size_hint=(1, None), height=80)

        layout.add_widget(self.refresh_button)
        layout.add_widget(self.search_input)
        layout.add_widget(self.status_filter)
        layout.add_widget(scroll_view)
        layout.add_widget(self.back_button)

        self.add_widget(layout)

    def get_credentials(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        credentials_file_path = os.path.join(script_dir, 'credentials.json')
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        return ServiceAccountCredentials.from_json_keyfile_name(credentials_file_path, scopes)

    def _refresh_codes(self, dt):
        # Change the button color to blue
        self.refresh_button.background_color = [0, 0, 1, 1]
        
        # Schedule the actual refreshing task to allow the UI to update
        Clock.schedule_once(lambda dt: self._perform_refresh_codes(dt), 0.1)

    def _perform_refresh_codes(self, dt):
        try:
            credentials = self.get_credentials()
            client = gspread.authorize(credentials)
            sheet = client.open('voucher-barcode').sheet1
            self.codes = sheet.get_all_values()[1:]  # Skip header row
            self.update_code_buttons()
        except Exception as e:
            self.show_error_popup("Error", str(e))
        finally:
            # Revert the button color to its original color
            self.refresh_button.background_color = [0.3, 0.7, 0.3, 1]


    def update_code_buttons(self, instance=None, value=''):
        self.code_list_layout.clear_widgets()
        search_text = self.search_input.text.lower()
        status_filter = self.status_filter.text
        filtered_codes = [code for code in self.codes if search_text in code[0].lower()]

        if status_filter != 'All':
            filtered_codes = [code for code in filtered_codes if self.get_code_status(code) == status_filter]

        for code_data in filtered_codes:
            code_button = Button(
                text=f"{code_data[0]} - {self.get_code_status(code_data)}",
                size_hint_y=None,
                height=60,
                font_size=16,
                on_press=lambda instance, cd=code_data: self.show_code_info(cd)
            )
            self.code_list_layout.add_widget(code_button)

    def get_code_status(self, code_data):
        if code_data[0].startswith('*'):
            return 'Unactive'
        elif code_data[0].startswith('~'):
            return 'Redeemed'
        else:
            return 'Active'

    def show_success_popup(self, title, message):
        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        label = Label(text=message)
        close_button = Button(text='Close', size_hint=(1, 0.2))
        close_button.bind(on_press=lambda x: popup.dismiss())
        layout.add_widget(label)
        layout.add_widget(close_button)
        popup = Popup(title=title, content=layout, size_hint=(0.8, 0.4))
        popup.open()

    def show_error_popup(self, title, message):
        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        label = Label(text=message)
        close_button = Button(text='Close', size_hint=(1, 0.2))
        close_button.bind(on_press=lambda x: popup.dismiss())
        layout.add_widget(label)
        layout.add_widget(close_button)
        popup = Popup(title=title, content=layout, size_hint=(0.8, 0.4))
        popup.open()

    def dismiss_popup(self, popup):
        popup.dismiss()

    def show_code_info(self, code_data):
        # Define the popup layout
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        # Determine button state based on code status
        activate_disabled = not code_data[0].startswith('*') or not code_data[2]
        redeem_disabled = code_data[0].startswith('*') or code_data[0].startswith('~')

        # Create buttons
        activate_button = Button(
            text='Activate', 
            on_press=lambda instance: self.activate_code(code_data[0], instance), 
            disabled=activate_disabled, 
            size_hint=(None, None), 
            size=(150, 40),
            font_size=14,
            background_color=[0.3, 0.7, 0.3, 1] if not activate_disabled else [0.7, 0.7, 0.7, 1]
        )
        redeem_button = Button(
            text='Redeem', 
            on_press=lambda instance: self.redeem_code(code_data[0], instance), 
            disabled=redeem_disabled, 
            size_hint=(None, None), 
            size=(150, 40),
            font_size=14,
            background_color=[0.3, 0.7, 0.3, 1] if not redeem_disabled else [0.7, 0.7, 0.7, 1]
        )

        # Create labels for code information
        status = 'Unactive' if code_data[0].startswith('*') else ('Redeemed' if code_data[0].startswith('~') else 'Active')
        info_label = Label(
            text=f"Code: {code_data[0]}\nStatus: {status}\nManier Verzonden: {code_data[2]}\nExpiry: {code_data[3]}",
            size_hint_y=None,
            height=150,
            font_size=21
        )

        # Add widgets to the popup content
        content.add_widget(info_label)
        content.add_widget(activate_button)
        content.add_widget(redeem_button)

        # Create and open the popup
        popup = Popup(title=f'Code Options: {code_data[0]}', content=content, size_hint=(None, None), size=(500, 400))
        popup.open()

    def activate_code(self, code, button):
        button.background_color = [0, 0, 1, 1]  # Change button to blue
        Clock.schedule_once(lambda dt: self._activate_code(code, button), 0.1)

    def _activate_code(self, code, button):
        try:
            # Simulate code activation
            print(f"Activating code: {code}")
            creds = self.get_credentials()
            client = gspread.authorize(creds)
            sheet = client.open('voucher-barcode').sheet1
            cell = sheet.find(code)
            if cell:
                if cell.value.startswith("*"):  # Check if the code is unactive
                    # Mark the code as active
                    sheet.update_cell(cell.row, 1, cell.value[1:])  # Remove the asterisk
                    # Set the expiry date to one year from today
                    expiry_date = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime("%m/%d/%Y")
                    sheet.update_cell(cell.row, 4, expiry_date)  # Assuming the expiry date is in the 4th column
                    print(f"Code activated: {code}")
                    self.show_success_popup("Success", f"Code activate, Expires on {expiry_date}.")
                else:
                    print("Code is already active or redeemed.")
                    self.show_error_popup("Error", "Code is already active or redeemed.")
            else:
                print("Code not found.")
                self.show_error_popup("Error", "Code not found.")
        except gspread.exceptions.SpreadsheetNotFound:
            print("Spreadsheet 'voucher-barcode' not found.")
            self.show_error_popup("Error", "Spreadsheet not found")
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            self.show_error_popup("Error", str(e))
        finally:
            button.background_color = [0.3, 0.7, 0.3, 1]  # Revert button color


    def redeem_code(self, code, button):
        button.background_color = [0, 0, 1, 1]
        Clock.schedule_once(lambda dt: self._redeem_code(code, button), 0.1)

    def _redeem_code(self, code, button):
        try:
            print(f"Redeeming code: {code}")
            creds = self.get_credentials()
            client = gspread.authorize(creds)
            sheet = client.open('voucher-barcode').sheet1
            cell = sheet.find(code)
            if cell:
                if not cell.value.startswith("~"):  # Check if the code is not already redeemed
                    # Mark the code as redeemed
                    sheet.update_cell(cell.row, 1, "~" + cell.value[1:])
                    # Get the current date formatted as MM/DD/YYYY
                    redemption_date = datetime.datetime.now().strftime("%m/%d/%Y")
                    # Update redemption date in column E (5th column)
                    sheet.update_cell(cell.row, 5, redemption_date)
                    print(f"Code redeemed. Valid until {redemption_date}.")
                    self.show_success_popup("Success", f"Code redeemed. Valid until {redemption_date}.")
                else:
                    print("Code is already redeemed.")
                    self.show_error_popup("Error", "Code is already redeemed.")
            else:
                print("Code not found.")
                self.show_error_popup("Error", "Code not found.")
        except gspread.exceptions.SpreadsheetNotFound:
            print("Spreadsheet 'voucher-barcode' not found.")
            self.show_error_popup("Error", "Spreadsheet not found")
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            self.show_error_popup("Error", str(e))
        finally:
            button.background_color = [1, 1, 1, 1]  # Revert button color
            
    def switch_to_main(self, instance):
        self.manager.current = 'main'

class ScrollViewApp(App):
    def build(self):
        # Create the main layout
        layout = BoxLayout(orientation='vertical')

        # Create a Label widget
        label = Label(text="Voucher Codes", size_hint_y=None, height=40)

        # Add the Label to the main layout
        layout.add_widget(label)

        # Create a ScrollView
        scroll_view = ScrollView(
            size_hint=(1, 1),  # Takes up the remaining height
            scroll_type=['bars', 'content'],  # Enable scrolling via bars and content dragging
            bar_width=10  # Width of the scroll bar
        )
        
        # Create a GridLayout that will be added to the ScrollView
        grid_layout = GridLayout(cols=1, spacing=10, size_hint_y=None)
        grid_layout.bind(minimum_height=grid_layout.setter('height'))
        
        # Add buttons to the GridLayout (replace with your voucher code buttons)
        for i in range(20):  # Example: 20 buttons for voucher codes
            btn = Button(text=f'Voucher Code {i + 1}', size_hint_y=None, height=40)
            grid_layout.add_widget(btn)
        
        # Add the GridLayout to the ScrollView
        scroll_view.add_widget(grid_layout)
        
        # Add the ScrollView to the main layout
        layout.add_widget(scroll_view)
        
        return layout

if __name__ == '__main__':
    VoucherApp().run()
