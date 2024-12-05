import win32com.client
import os
import time
import pandas as pd

psApp = win32com.client.Dispatch("Photoshop.Application")
os.startfile("PhotoXcell.psd")


# Load Excel file
excel_file = 'PhotoXcell.xlsx'
data = pd.read_excel(excel_file)

output_dir = 'output_images'
os.makedirs(output_dir, exist_ok=True)
print("Waiting for Photoshop to Open.")
print("!!! if Photoshop Open's After 5seconds, re-run the code.")
time.sleep(5) 
doc = psApp.Application.ActiveDocument

for index, row in data.iterrows():
        user_code = row['Code']  # Replace with your column Code
        user_name = row['Name']  # Replace with your column Name
        layers = doc.ArtLayers["UserCode"]
        text_of_layer = layers.TextItem
        
        text_of_layer.contents = str(user_code)

        options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
        options.Format = 13   # PNG Format
        options.PNG8 = False  # Sets it to PNG-24 bit


        output_file = f'{os.path.dirname(__file__)}/{output_dir}/{user_name}.png'
        doc.Export(ExportIn=output_file, ExportAs=2, Options=options)
        print(f'Saved: {output_file}')

