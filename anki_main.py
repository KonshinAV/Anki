from deep_translator import GoogleTranslator
from pprint import pprint
import json
import uuid
import requests
import openpyxl

class SimpleGoogleTranslate:
    def __init__(self, source='de', target='ru'):
        self.source = source
        self.target = target

    def translate(self, text):
        return GoogleTranslator(source=self.source, target=self.target).translate(text=text)

class Anki:
    def __init__(self, ANKI_CONNECT_URL = 'http://localhost:8765'):
        self.ANKI_CONNECT_URL = ANKI_CONNECT_URL
        pass

    def invoke(self, action, params=None):
        response = requests.post(self.ANKI_CONNECT_URL, json={
            'action': action,
            'version': 6,
            'params': params or {}
        }).json()
        if response.get("error") is not None:
            raise Exception(response["error"])
        return response["result"]

    def get_note_ids_from_deck(self, deck_name):
        query = f'deck:"{deck_name}"'
        return self.invoke('findNotes', {'query': query})

    def get_notes_info(self, note_ids):
        return self.invoke('notesInfo', {'notes': note_ids})

    def show_notes_from_deck(self, deck_name):
        note_ids = self.get_note_ids_from_deck(deck_name)
        print(f"Найдено заметок в колоде '{deck_name}': {len(note_ids)}")
        
        if not note_ids:
            print("Нет заметок.")
            return
        
        notes = self.get_notes_info(note_ids)
        for note in notes:
            pprint(note)

    def show_note_by_id(self, note_id):
        result = self.get_notes_info([note_id])
        if not result:
            print(f"Заметка с ID {note_id} не найдена.")
            return
        note = result[0]
        pprint(note)

    def update_note_field(self, note_id, field_name, new_value):
        note_info = self.get_notes_info([note_id])
        if not note_info:
            print(f"Заметка с ID {note_id} не найдена.")
            return

        # Проверка существования поля
        fields = note_info[0]['fields']
        if field_name not in fields:
            print(f"Поле '{field_name}' не найдено в заметке.")
            return

        # Обновление поля
        self.invoke('updateNoteFields', {
            'note': {
                'id': note_id,
                'fields': {
                    field_name: new_value
                }
            }
        })
        print(f"Поле '{field_name}' обновлено на: {new_value}")

    def translate_base(self, notes, de_field, ru_field, update_only_empty_values = True):
        notes_count = len(notes)
        for ind,note in enumerate(notes):
            note_id = note['noteId']
            de_value = note['fields'][de_field]['value']
            ru_value = note['fields'][ru_field]['value']
            print(f"{ind+1} von {notes_count}, {de_field}: {de_value}")
            if de_value != '':
                if update_only_empty_values:
                    if ru_value == '':
                        self.update_note_field(note_id=note_id,
                                            field_name=ru_field,
                                            new_value=translator.translate(text=de_value))
                    else:
                        print(f"Поле {ru_field} не пустое и будет пропущено. Текущее значение: {ru_value}")
                else:
                    self.update_note_field(note_id=note_id,
                                            field_name=ru_field,
                                            new_value=translator.translate(text=de_value))
            else:
                print(f"Field {de_field} is empty")

    def generate_and_insert_tts(self, notes, source_field, target_field):
        for note in notes:
            note_id = note['noteId']
            note_info = self.get_notes_info([note_id])
            if not note_info:
                print(f"Заметка с ID {note_id} не найдена.")
                return

            note = note_info[0]
            fields = note['fields']
            if source_field not in fields:
                print(f"Поле '{source_field}' не найдено.")
                return

            text = fields[source_field]['value']
            if not text.strip():
                print(f"Поле '{source_field}' пустое.")
                return

            # Генерация уникального имени файла
            uid = str(uuid.uuid4())  # Убираем дефисы для компактности
            filename = f"google-{uid}.mp3"

            print(f"Генерация озвучки для текста: {text}")
            print(f"Имя файла: {filename}")

            # Генерация аудио-файла
            self.invoke("ttsMake", {
                        "text": text,
                        "voice": "GoogleTranslate:de",
                        "cache": True,
                        "slow": False,
                        "useCacheOnly": False,
                        "filename": filename
                    })

            # Вставка [sound:filename] в нужное поле
            sound_tag = f"[sound:{filename}]"

            result = invoke("updateNoteFields", {
                "note": {
                    "id": note_id,
                    "fields": {
                        target_field: sound_tag
                    }
                }
            })

            # Добавление самого аудиофайла в заметку
            self.invoke("addNoteAudio", {
                "note": {
                    "id": note_id
                },
                "audio": {
                    "filename": filename,
                    "fields": [target_field],
                    "skipHash": "7e2c2f954ef6051373ba916f000168dc"
                }
            })

def read_xlsx_file(filename, sheet_name):
    # Загружаем файл
    workbook = openpyxl.load_workbook(filename)

    # Проверяем, существует ли лист с указанным именем
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Лист с именем '{sheet_name}' не найден в файле.")

    # Получаем нужный лист
    sheet = workbook[sheet_name]

    # Получаем заголовки из первой строки
    header = [cell for cell in next(sheet.iter_rows(values_only=True))]

    # Читаем остальные строки и формируем список словарей, пропуская пустые строки
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if any(cell is not None for cell in row):  # Проверка на непустую строку
            row_dict = dict(zip(header, row))
            data.append(row_dict)

    return data

if __name__ == "__main__":
    translator = SimpleGoogleTranslate()
    anki = Anki()
    # deck_name = "Deutsche Lernen::B1_Wortliste_lernen"
    deck_name = "Deutsche Lernen::Dev_deck"
    
    anki_note_ids = anki.get_note_ids_from_deck(deck_name=deck_name)
    anki_notes = anki.get_notes_info(note_ids=anki_note_ids)
    pprint(anki.show_notes_from_deck(deck_name=deck_name))
    # anki.generate_and_insert_tts(notes=anki_notes,source_field="audio_text_de",target_field="base_audio")
    # anki.show_notes_from_deck(deck_name=deck_name)

    # anki.translate_base(notes=anki_notes, de_field='base_de', ru_field='base_ru')
    # anki.translate_base(notes=anki_notes, de_field='s3_de', ru_field='s3_ru')

    # xlsx_file_content = read_xlsx_file(filename="for_import.xlsx", sheet_name="Sheet2")
    # pprint(xlsx_file_content)
