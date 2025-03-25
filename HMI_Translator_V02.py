

def Translate(File_Path,API_Key,Rows,Language,Startrow):

 import openpyxl
 import requests
 import time

 from openpyxl import Workbook, load_workbook
 book = load_workbook(File_Path)
 sheet = book.active
 

 def translate_word(word, api_key=API_Key):
     url = "https://api-free.deepl.com/v2/translate"
     payload = {
        'auth_key': api_key,
        'text': word,
        'target_lang': Language
     }
     response = requests.post(url, data=payload)
     translation = response.json()['translations'][0]['text']
     return translation



    # Translate the word


 for i in range(Startrow, Rows):  # Starting from 1 since Excel rows are 1-indexed, and going up to 6 to include the 5th row
     start_time = time.time()
     end_time = time.time()
     cell_value = sheet[f'F{i}'].value  # Accessing the cell value in column A and row i
     if (cell_value is not None) and \
   ("<field ref=\"2\" />" not in cell_value) and \
   ("<field ref=\"0\" />" not in cell_value) and \
   ("<field ref=\"1\" />" not in cell_value):
       translated_word = translate_word(cell_value)
       sheet[f'G{i}'].value = translated_word
       end_time = time.time()  # Capture end time of iteration
     iteration_duration = end_time - start_time  # Calculate duration of iteration
     if i % 100 == 0:  # Save after every 100 translations
                    print(f"Saving progress at row {i}...")
                    book.save(File_Path)
     print('Line N°',i,'Done in ',f"{iteration_duration:.3f}",'s')

 book.save(File_Path)


 print ('Translation done')
 return (i)
