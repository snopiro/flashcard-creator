import docx
import requests
import os
import tkinter as tk
import sys
from tkinter import filedialog, simpledialog, messagebox, ttk


def show_instructions_and_confirm():
    """Display instructions and wait for the user's confirmation to proceed."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    proceed = messagebox.askokcancel("Instructions", 
                                     "This script will allow you to import flashcards from a Word document into Anki.\n\n"
                                     "1. You'll first be prompted to select a .docx file.\n"
                                     "2. Next, you'll specify the name of the Anki deck to add cards to.\n\n"
                                     "Ensure Anki is running with AnkiConnect installed.\n\n"
                                     "To install AnkiConnect, Open Anki, go to Tools->Addons->Get Addons\n"
                                     "and enter the code 2055492159 \n\n"
                                     "Click OK to continue or Cancel to exit.")
    if not proceed:
        print("User opted to exit. Exiting.")
        sys.exit()

def extract_flashcards_from_docx(document_path):
    """Extracts table data from a Word document and returns it as a list of flashcards."""
    # Load the Word document
    doc = docx.Document(document_path)

    # Assuming the table is the first in the document
    table = doc.tables[0]

    flashcards = []

    # Iterate through rows in the table (skip the header row)
    for row in table.rows[1:]:
        vocabulary, kanji, translation = [cell.text for cell in row.cells]
        # Wrap kanji with parentheses if it exists
        kanji = f"({kanji.strip()})" if kanji.strip() else ""
        flashcards.append({
            'vocabulary': vocabulary.strip(),
            'kanji': kanji,
            'translation': translation.strip()
        })
    return flashcards

def ensure_deck_exists(deck_name, anki_connect_url):
    """Ensure the deck exists, or create it if it doesn't."""
    payload = {
        "action": "deckNames",
        "version": 6
    }
    response = requests.post(anki_connect_url, json=payload)
    existing_decks = response.json().get('result', [])

    if deck_name not in existing_decks:
        payload = {
            "action": "createDeck",
            "version": 6,
            "params": {
                "deck": deck_name
            }
        }
        requests.post(anki_connect_url, json=payload)

def prompt_for_filepath():
    """Use a file dialog to get a valid docx file path."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title="Select a Word document", filetypes=[("Word Documents", "*.docx")])
    if file_path:
        return file_path
    else:
        print("No file selected. Exiting.")
        sys.exit()

def prompt_for_deck_name():
    """Use a dialog to get the deck name."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    deck_name = simpledialog.askstring("Input", "Please specify the name of the Anki deck you want to add cards to:")
    if deck_name:
        return deck_name.strip()
    else:
        print("No deck name provided. Exiting.")
        sys.exit()
def add_flashcards_to_anki(flashcards, deck_name):
    """Add flashcards to Anki and show progress using a GUI progress bar."""
    
    root = tk.Tk()
    root.title('Adding Flashcards...')

    # Create and layout the progress bar
    progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
    progress.pack(pady=20)
    label = tk.Label(root, text="Adding Flashcards...")
    label.pack(pady=10)

    progress["maximum"] = len(flashcards)
    progress["value"] = 0
    root.update()

    anki_connect_url = "http://localhost:8765"
    for flashcard in flashcards:
        note = {
            "action": "addNote",  # Specify the action
            "version": 6,  # AnkiConnect version
            "params": {
                "note": {
                    "deckName": deck_name,
                    "modelName": "Basic",
                    "fields": {
                        "Front": f"{flashcard['vocabulary']} {flashcard['kanji']}",
                        "Back": flashcard['translation']
                    },
                    "options": {
                        "allowDuplicate": False
                    }
                }
            }
        }


        # Send the note to Anki using AnkiConnect
        response = requests.post(f"{anki_connect_url}/addNote", json=note)
        response_data = response.json()
        if response_data.get('error'):
            print(f"Failed to add flashcard: {flashcard['vocabulary']}. Error: {response_data.get('error')}")

        # Update progress bar
        print(f"Added flashcard: {flashcard['vocabulary']}") # Print to console
        progress["value"] += 1
        root.update()

    # Once all flashcards are added, destroy the progress bar window
    root.destroy()


if __name__ == "__main__":
    show_instructions_and_confirm()

    document_path = prompt_for_filepath()
    deck_name = prompt_for_deck_name()

    flashcards = extract_flashcards_from_docx(document_path)

    # AnkiConnect API endpoint
    anki_connect_url = "http://localhost:8765"

    # Ensure the deck exists
    ensure_deck_exists(deck_name, anki_connect_url)

    # Iterate through flashcards and add them to Anki
    add_flashcards_to_anki(flashcards, deck_name)


    # Show a finished confirmation box after all flashcards are added
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showinfo("Completed", "All flashcards have been successfully added to Anki!")