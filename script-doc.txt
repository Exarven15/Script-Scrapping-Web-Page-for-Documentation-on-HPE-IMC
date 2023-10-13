import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# URL de la page web que vous souhaitez scraper
url = input("Entrez l'URL de la page : ")

try:
    # Envoyez une requête HTTP GET à l'URL
    response = requests.get(url)

    # Vérifiez si la requête a réussi (code de statut 200)
    if response.status_code == 200:
        # Analysez le contenu HTML de la page à l'aide de BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')

        # Recherchez l'élément h1 avec la classe PageTitle
        title_element = soup.find('h1', class_='PageTitle')

        # Vérifiez si l'élément a été trouvé
        if title_element:
            # Récupérez le texte du titre
            title_text = title_element.get_text()

            # Recherchez l'élément avec la classe spécifiée
            target_element = soup.find('div', class_='lia-quilt-column lia-quilt-column-24 lia-quilt-column-single lia-quilt-column-main-content')

            # Vérifiez si l'élément cible a été trouvé
            if target_element:
                # Récupérez le texte de l'élément cible
                target_text = target_element.get_text()

                # Créez un nom de fichier unique basé sur le titre
                filename = f'{title_text}.txt'
                excel_filename = 'annuaire-doc.xlsx'

                # Enregistrez le contenu dans le fichier texte
                with open(filename, 'w', encoding='utf-8') as file:
                    file.write(target_text)
                print(f'Contenu enregistré dans le fichier {filename}')

                # Chargez le fichier Excel existant ou créez-en un nouveau s'il n'existe pas
                try:
                    workbook = load_workbook(excel_filename)
                    worksheet = workbook.active
                except FileNotFoundError:
                    workbook = Workbook()
                    worksheet = workbook.active
                    worksheet.append(['Lien vers le fichier texte'])
                    worksheet['A1'].font = Font(underline="single", color="0000FF")

                # Ajoutez un lien vers le fichier texte dans une nouvelle ligne du fichier Excel
                new_row = [f'=HYPERLINK("{filename}", "{title_text}")']
                worksheet.append(new_row)

                # Enregistrez le fichier Excel avec la nouvelle ligne ajoutée
                workbook.save(excel_filename)
                print(f'Fichier Excel mis à jour avec le lien sous {excel_filename}')
            else:
                print("L'élément cible n'a pas été trouvé sur la page.")
        else:
            print("L'élément de titre avec la classe PageTitle n'a pas été trouvé sur la page.")
    else:
        print(f'Impossible de récupérer la page. Code de statut : {response.status_code}')

except Exception as e:
    print(f'Une erreur s'est produite : {str(e)}')