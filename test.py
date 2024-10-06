import os

path_to_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# Test pour vérifier si wkhtmltopdf est accessible
if not os.path.exists(path_to_wkhtmltopdf):
    print(f"wkhtmltopdf non trouvé à l'emplacement: {path_to_wkhtmltopdf}")
else:
    print(f"wkhtmltopdf trouvé à: {path_to_wkhtmltopdf}")
