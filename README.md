# Voucher Barcode Generator en Manager

Deze repository bevat een Python-script dat unieke vouchercodes genereert, opslaat als barcodes en hun status beheert (bijv. activatie, inwisseling). Het integreert ook met Google Sheets voor het opslaan en beheren van de vouchercodes.

Dit script is uitsluitend bedoeld voor medewerkers van Fietsstation en de eigenaar.

## Inhoudsopgave

- [Vereisten](#vereisten)
- [Installatie](#installatie)
- [Gebruik](#gebruik)
- [Hoofdmenu](#Hoofdmenu)
- [Output File](#output-file)
- [Default Folder](#default-folder)
- [Spreadsheet Toegang](#spreadsheet-toegang)
- [Functies](#functies)

## Vereisten

Voordat je begint, zorg ervoor dat je aan de volgende vereisten voldoet:

- Internettoegang om de Google Sheets API te gebruiken.

## Installatie

1. Download het zip-bestand van de repository.
2. Pak het zip-bestand uit naar een gewenste locatie op je computer.
3. Open de uitgepakte map en voer de setup-folder uit.
4. Volg de stappen in de folder:
    - **Stap 1**: Installatie van Python 3.11  
      Open het bestand genaamd `step 1 Python 3.11.cmd`
    - **Stap 2**: Installatie van de vereiste pakketten  
      Open het bestand genaamd `step 2 Pip install.cmd`

Je hoeft geen `credentials.json` bestand te plaatsen; dit is al inbegrepen in de download.

## Gebruik

Klik op het `voucher.py` bestand om het script te starten.

Je zult 2 versies zien
`Voucher-EG.py` En `Voucher-NL.py`, 
`Voucher-EG.py` Is de engelse versie en `Voucher-NL.py` is de nederlandse versie.
Je kunt zelf kiezen welke je wilt gebruiken
Het script zal je een menu presenteren om verschillende bewerkingen uit te voeren, zoals het genereren van codes, het controleren van de geldigheid van een code, het afdrukken van codes, het bekijken van informatie, het activeren en inwisselen van codes.

## Hoofdmenu

Deze uitleg moet nog geschreven worden

## Output File

De gegenereerde coupons en barcodes worden opgeslagen in een output-bestand. Dit bestand bevat alle uitkomsten van het script, zoals de gegenereerde barcodes en coupons. Zorg ervoor dat je de locatie van dit bestand noteert, zodat je eenvoudig toegang hebt tot de opgeslagen gegevens.

## Default Folder

De folder genaamd `Default` bevat alle standaard ontwerpen die het script nodig heeft om de coupons en barcodes te maken. Zorg ervoor dat deze folder aanwezig is in de hoofdmap van het project, omdat het script anders niet correct zal functioneren.

## Spreadsheet Toegang

In de hoofdmap van de uitgepakte bestanden bevindt zich een bestand genaamd `Open_spreadsheet.cmd`. Dit bestand opent een link naar de Google Spreadsheet die alle informatie en codes bevat. Door dit bestand uit te voeren, krijg je eenvoudig toegang tot de spreadsheet waar je de status en details van alle vouchercodes kunt beheren.

## Functies

### `clear_screen()`
Maakt het terminalscherm leeg.

### `loading_bar(action, *args)`
Toont een laadbalk tijdens het uitvoeren van de opgegeven actie.

### `generate_unique_code()`
Genereert een unieke alfanumerieke code.

### `generate_barcode(code, filename)`
Genereert een barcode-afbeelding voor de opgegeven code.

### `get_credentials(credentials_file_path=None)`
Haalt Google Sheets API-referenties op.

### `add_to_google_sheet(num_codes, credentials_file_path=CREDENTIALS_FILE_PATH)`
Genereert en voegt unieke codes toe aan een Google Sheet.

### `activate_code(code, credentials_file_path=None)`
Activeert een opgegeven code in de Google Sheet.

### `check_code_validity(code, credentials_file_path=CREDENTIALS_FILE_PATH)`
Controleert de geldigheid en status van een opgegeven code.

### `print_image_in_portrait(file_path)`
Drukt een afbeeldingsbestand af in portretmodus.

### `save_coupon(code)`
Genereert een couponafbeelding met de opgegeven code en barcode.

### `make_printed()`
Markeert geselecteerde niet-geprinte codes als geprint en genereert een PDF van de coupons.

### `find_random_unprinted_code(credentials_file_path=None)`
Zoekt een willekeurige niet-geprinte code uit de Google Sheet.

### `mark_as_printed(code, credentials_file_path=CREDENTIALS_FILE_PATH)`
Markeert een opgegeven code als geprint in de Google Sheet.

### `print_images_in_windows(file_paths)`
Drukt afbeeldingen af met de standaardprinter op Windows.

### `info_status()`
Toont het statusmenu voor verschillende informatie over codes.

### `show_basic_info()`
Toont basisinformatie over de codes.

### `show_all_codes()`
Toont alle codes en hun status.

### `redeem_code(code)`
Wisselt een opgegeven code in door deze als gebruikt te markeren in de Google Sheet.

### `main()`
De hoofdfunctie die een gebruikersinterface biedt voor het beheren van vouchercodes.

---

**Gemaakt door**: Milan Vosters  
**Project gestart op**: 08:00AM 16/5/2024

Voor eventuele problemen of vragen, open een issue op GitHub of neem contact met mij op via [m.vosters@beginstation.nl](mailto:m.vosters@beginstation.nl).
