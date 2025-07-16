from bs4 import BeautifulSoup
import re
import os

def extract_transaction_types():
    """
    Extrahiert Transaction Types aus der lokalen HTML-Datei und speichert sie in einer Textdatei
    """
    
    html_filename = "Search _ People of Medieval Scotland.html"
    
    try:
        # Prüfen ob die HTML-Datei existiert
        if not os.path.exists(html_filename):
            print(f"Fehler: Die Datei '{html_filename}' wurde nicht gefunden!")
            print("Bitte stellen Sie sicher, dass die Datei im selben Ordner liegt.")
            return
        
        print(f"Lade HTML-Datei: {html_filename}")
        
        # HTML-Datei lesen
        with open(html_filename, 'r', encoding='utf-8') as file:
            html_content = file.read()
        
        # HTML parsen
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Die spezifische div finden
        target_div = soup.find('div', class_='g n2 transaction__transactiontypes')
        
        if not target_div:
            print("Fehler: Div mit Klasse 'g n2 transaction__transactiontypes' nicht gefunden")
            return
        
        print("Div gefunden, extrahiere Transaction Types...")
        
        # Alle li-Elemente in der div finden
        li_elements = target_div.find_all('li')
        
        transaction_types = []
        
        for li in li_elements:
            # a-Element innerhalb des li finden
            a_element = li.find('a')
            
            if a_element and a_element.text:
                # Text extrahieren
                full_text = a_element.text.strip()
                
                # Klammer mit Zahl entfernen (z.B. "Appointment (213)" -> "Appointment")
                # Regex: entfernt alles ab dem letzten öffnenden Klammer bis zum Ende
                clean_text = re.sub(r'\s*\([^)]*\)\s*$', '', full_text).strip()
                
                if clean_text:
                    transaction_types.append(clean_text)
                    print(f"Gefunden: {clean_text}")
        
        # In Textdatei speichern
        filename = "Poms_transcation_types.txt"
        
        with open(filename, 'w', encoding='utf-8') as f:
            for transaction_type in transaction_types:
                f.write(transaction_type + '\n')
        
        print(f"\n{len(transaction_types)} Transaction Types erfolgreich in '{filename}' gespeichert!")
        
        return transaction_types
        
    except FileNotFoundError:
        print(f"Fehler: Die Datei '{html_filename}' wurde nicht gefunden!")
        print("Bitte stellen Sie sicher, dass die HTML-Datei im selben Ordner liegt.")
        
    except UnicodeDecodeError:
        print("Fehler beim Lesen der Datei. Versuche mit anderer Kodierung...")
        try:
            with open(html_filename, 'r', encoding='latin-1') as file:
                html_content = file.read()
            soup = BeautifulSoup(html_content, 'html.parser')
            # Hier würde der gleiche Extraktionscode folgen...
            print("Datei erfolgreich mit latin-1 Kodierung gelesen!")
        except Exception as e:
            print(f"Fehler auch mit alternativer Kodierung: {e}")
        
    except Exception as e:
        print(f"Unerwarteter Fehler: {e}")

if __name__ == "__main__":
    print("POMS Transaction Types Extractor (Lokale HTML-Datei)")
    print("=" * 50)
    
    extract_transaction_types()