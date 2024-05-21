# Voucher code generator and coupon maker

Deze repository bevat een Python-script dat unieke vouchercodes genereert, opslaat als barcodes en hun status beheert (bijv. activatie, inwisseling). Het integreert ook met Google Sheets voor het opslaan en beheren van de vouchercodes.

Dit script is uitsluitend bedoeld voor medewerkers van Fietsstation en de eigenaar.

## Inhoudsopgave

- [Vereisten](#vereisten)
- [Installatie](#installatie)
- [Gebruik](#gebruik)
- [Hoofd Menu](#hoofd-menu)
   - [Optie 1](#optie-1)
   - [Optie 2](#optie-2)
   - [Optie 3](#optie-3)
   - [Optie 4](#optie-4)
   - [Optie 5](#optie-5)
   - [Optie 6](#optie-6)
   - [Optie 7](#optie-7)
- [Output File](#output-file)
- [Default Folder](#default-folder)
- [Spreadsheet Toegang](#spreadsheet-toegang)
- [Functies](#functies)
- [Errors](#errors)
- [Contact](#contact)

## Vereisten

Voordat je begint, zorg ervoor dat je aan de volgende vereisten voldoet:

- Internettoegang om de Google Sheets API te gebruiken.

## Installatie

1. Download het zip-bestand van de repository.
2. Pak het zip-bestand uit naar een gewenste locatie op je computer.
3. Open de uitgepakte map en voer de setup-folder uit.
4. Volg de stappen in de folder:
    - **Stap 1**: Installatie van Python  
      Open het bestand genaamd `step 1 Python.cmd`
      Als windows zegt dat het bestand onveilig is klik op *Meer info/More info*
      En klik daarna op Run<br />
![](Data/gifs/step1.gif)
      Download python via de site die wordt geopent
    - **Stap 2**: Installatie van de vereiste pakketten  
      Open het bestand genaamd `step 2 Pip install.cmd`
      Dit instaleerd alle modules je hoeft niks te doen terwijl dit laad
      Het kan tot 2min max duren.
      Ook Dit script kan voor de eerste keer zeggen dat het bestand onveilig is, doe het zelfde om hier omheen te komen als stap 1.<br />
![](Data/gifs/step2.gif)
Je hoeft geen `credentials.json` bestand te plaatsen; dit is al inbegrepen in de download.

## Gebruik

Klik op het `voucher.py` bestand om het script te starten.

Je zult 2 versies zien
`Voucher-EG.py` En `Voucher-NL.py`, 
`Voucher-EG.py` Is de engelse versie en `Voucher-NL.py` is de nederlandse versie.
Je kunt zelf kiezen welke je wilt gebruiken
Het script zal je een menu presenteren om verschillende bewerkingen uit te voeren, zoals het genereren van codes, het controleren van de geldigheid van een code, het afdrukken van codes, het bekijken van informatie, het activeren en inwisselen van codes.

## Hoofd menu

  ### Menu 
  Dit is wat je ziet wanneer je het programma start.<br />
![](Data/picture/main.png)<br />
  ### Optie 1
   - Deze optie is `Code Generator/Code Genarator`
Met deze optie kun je codes genereren in het spreadsheet `Open_spreadsheet.cmd`.
Je kunt een aantal invoeren en deze zullen gegenereerd worden binnen 10 seconden.<br />
![](Data/picture/generator.png)<br />
  ### Optie 2
  Deze optie is `Code verifieren/Verify Code`
   Met deze optie kun je een code van een klant of `Open_spreadsheet.cmd` pakken en checken of hij geldig is en wat de status is.<br />
![](Data/picture/verify.png)<br />
  ### Optie 3
Dit is de optie `Print/Printen`
   Met deze optie worden er **36** Random ongeactiveerde codes gevonden en op een blaadje met de coupons als barcode geplaatst.<br />
Dit is een voorbeeld:
 [Coupon_List.pdf](https://github.com/BeginstationMilan/Voucher-Program/files/15388530/Coupon_List.pdf)
![](Data/picture/print.png)<br />
Als de maat niet klopt of er is iets fout ga naar [Contact](#contact) onderaan en stuur een email.

  ### Optie 4
Dit is een mini menu voor info<br />
   ![](Data/picture/info.png)<br />
  ### Optie 4 1
Dit is Basis info dit geeft een lijst van het volgende: <br />
`Totaal aantal codes` = Aantal codes <br />
`Actieve codes` = Aantal actieve codes <br />
`Inactieve codes` = Aantal inactieve codes <br />
`Geprinte codes` = Aantal geprinte codes <br />
`Gebruikte codes` = Aantal gebruikte codes <br />
`Programma gemaakt door: Milan Vosters` = Info over maker <br />
`Project gestart op: 08:00AM 16/5/2024` = Info over Project <br />
      
  ### Optie 4 2
Dit is de `Alle codes` optie.
      Dit laat alle codes zien en hun status<br />
![](Data/picture/alle.png)<br />
  ### Optie 5
Dit is de optie `Code activeren/Activate code`
   Deze optie is om een code te activeren.
Dit doe je wanneer je een klant een coupon gaat geven.
Anders kunnen ze het niet inleveren.  <br />
Nadat je het hebt geactiveerd zet je een vinkje op geactiveerd op de coupon.<br />
![](Data/picture/activeer.png)<br />
  ### Optie 6
Deze optie is `Lever code in/Redeem code`
   Deze optie is om een code in te leveren van een klant.
Dit maakt de code ongeldig en dan krijgt de klant 10% korting<br />
![](Data/picture/lever.png)<br />
  ### Optie 7
Dit is de optie `Sluit af/Exit` 
   Dit sluit het programma af.<br />


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

# Errors

Hier zijn een paar mogelijke errors

   - **Programma start niet**<br />
dubbel check of je stap 1 `step 1 Python 3.11.cmd` & 2 `step 2 Pip install.cmd` hebt gedaan<br />
Als je stap 1 niet doet heeft het programma geen software om te runnen <br />
Als je stap 2 niet doet dan heeft het script niet de modules om te kunnen werken

   - **Optie 3 `Niet genoeg onactieve codes`** <br />
   Dit betekent dat het script niet genoeg onactieve ongeprinte codes heeft kunnen vinden<br />
   Je kunt dit oplossen om meer codes te genereren<br />
   Het script heeft er **36** nodig<br />

   - **Optie 2 `Code niet gevonden`** <br />
   Check of je de code goed hebt opgeschreven<br />
   Het kan soms irritant zijn met kleine letter L en hoofdletter i Il

   
---
## Contact

**Gemaakt door**: Milan Vosters  
**Project gestart op**: 08:00AM 16/5/2024

Voor eventuele problemen of vragen, open een issue op GitHub of neem contact met mij op via [m.vosters@beginstation.nl](mailto:m.vosters@beginstation.nl).
