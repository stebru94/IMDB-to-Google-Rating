from openpyxl import workbook
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import re
import time

FILM_DA_CERCARE = []
FILM_CLASSIFICATI = []
contatore_rimanenti = 1

def premiPulsanteAltri50():
    # Trova il pulsante "Altri 50" e il contenitore dei risultati
    buttonMore = driver.find_element(By.XPATH, "//span[@class='ipc-see-more__text' and text()='Altri 50']")
    results_container = driver.find_element(By.CLASS_NAME, 'ipc-title__text')

    # Scroll fino al pulsante "Altri 50"
    driver.execute_script("arguments[0].scrollIntoView();", buttonMore)
    time.sleep(1)

    # Fai clic sul pulsante "Altri 50"
    buttonMore.click()
    time.sleep(3)

def pausa():
    print("premi invio per continuare...")
    input()

def ottieniTitoli():
    # Apri la pagina web
    driver.get("https://www.imdb.com/search/title/?title_type=feature&release_date=1985-01-01,2023-12-31&genres=thriller,horror,crime")

    # Attendi che la pagina sia caricata
    driver.implicitly_wait(1)

    # clicca su accetto tutti i cookie
    try:
        button = driver.find_element(By.XPATH, '//*[@id="__next"]/div/div/div[2]/div/button[2]')
        button.click()
    except:
        pass

    pausa()

    #premiPulsanteAltri50()
    #premiPulsanteAltri50()



    # Trova il contenitore dei risultati
    content = driver.find_elements(By.CLASS_NAME, 'ipc-title__text')

    # Estrai i dati per ogni risultato e formatta il titolo
    for titolo_element in content:
        titolo = titolo_element.text
        titolo_senza_numero = re.sub(r'^\d+\.\s*', '', titolo)
        FILM_DA_CERCARE.append(titolo_senza_numero)

    # Stampa i titoli aggiunti alla lista
    for titolo in FILM_DA_CERCARE:
        print(titolo)

    print("film trovati: ", len(FILM_DA_CERCARE))


def formatta_titolo(titolo):
    # Sostituisci gli spazi con il segno più (+)
    titolo_formattato = titolo.replace(" ", "+")
    titolo_formattato = titolo_formattato + "+film"
    return titolo_formattato

def ricercaFilm(film):
    # Apri la pagina web
    driver.get("https://www.google.com/search?q=" + formatta_titolo(film))

    # Attendi che la pagina sia caricata
    driver.implicitly_wait(1)

    # clicca su accetto tutti i cookie
    try:
        button = driver.find_element(By.ID, "L2AGLb")
        button.click()
    except:
        pass

    # Attendi fino a 10 secondi per la presenza dell'elemento con la classe ".a19vA"
    try:
        content = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".a19vA"))
        )

        # Estrai i dati per ogni risultato
        punteggio_text = content.text

        punteggio_match = re.search(r'\d+%', punteggio_text)

        if punteggio_match:
            punteggio = punteggio_match.group()
            FILM_CLASSIFICATI.append((film, punteggio))
            contatore_rimanenti = len(FILM_DA_CERCARE) - len(FILM_CLASSIFICATI)
            print(f"{contatore_rimanenti} film rimanenti")

        else:
            print(f"{film}: N.D.")

    except TimeoutException:
        print(f"{film}: N.D.")

# Carica il file Excel e leggi i film già presenti
existing_films = set()
try:
    workbook = load_workbook('film_classificati.xlsx')
    sheet = workbook.active
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
        existing_films.add(row[0])
except FileNotFoundError:
    pass  # Se il file non esiste, non ci sono film già presenti

# Imposta il path del driver
driver_path = "C:/Chromedriver/chromedriver.exe"

# Configura le opzioni del driver
options = webdriver.ChromeOptions()

# Inizializza il driver di Chrome con le opzioni
driver = webdriver.Chrome(service=Service(driver_path), options=options)



ottieniTitoli()

for film in FILM_DA_CERCARE:
    if film in existing_films:
        print(f"Il film '{film}' è già presente nel database. Saltando la ricerca su Google.")
        continue
    ricercaFilm(film)

# Ordina gli elementi in base al punteggio (in ordine decrescente)
FILM_CLASSIFICATI.sort(key=lambda x: int(x[1].rstrip('%')), reverse=True)

# Stampa gli elementi ordinati
for film, punteggio in FILM_CLASSIFICATI:
    print(f"{film}: {punteggio}")

# Chiudi il driver
driver.quit()

# Verifica se il file Excel esiste già
try:
    workbook = load_workbook('film_classificati.xlsx')
    sheet = workbook.active
except FileNotFoundError:
    workbook = workbook()
    sheet = workbook.active
    sheet['A1'] = 'Titolo'
    sheet['B1'] = 'Percentuale'

# Aggiungi i dati dei film al foglio Excel solo se non sono già presenti
for film, punteggio in FILM_CLASSIFICATI:
    film_presente = False
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
        if film in row:
            film_presente = True
            break
    if not film_presente:
        sheet.append([film, punteggio])

# Salva il foglio Excel
workbook.save('film_classificati.xlsx')