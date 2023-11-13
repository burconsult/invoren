# InvoRen PDF Invoice Files Renamer
A Python desktop APP for renaming PDF files. This Python script can be packed as a desktop app. By dropping PDF files (invoices) onto the main window, information is extracted (invoice number from 1001 to 1099 in this version).
The PDF file is then renamed based on this information and the current date.
# Desktop app packaging
Use Pyinstaller to package it as an executable. Tkinterdnd2 (the drag and drop module) may cause issues so use the following command.
```pyinstaller --noconfirm --onefile --windowed --add-data "<path to Python>\Lib\site-packages\tkinterdnd2;tkinterdnd2\" .\invapp.py```

