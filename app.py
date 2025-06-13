from flask import Flask, render_template, request, redirect
from openpyxl import Workbook, load_workbook
import os
from datetime import datetime

app = Flask(__name__)
EXCEL_FILE = "reponses.xlsx"

# Création du fichier Excel s’il n’existe pas
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.append([
        "Date", "Nom", "Prénom", "Adresse", "Ville", "Code postal",
        "Rue ou résidence", "Ancienneté", "Ce que j'aime", "À améliorer",
        "Activité souhaitée", "Type d’événements", "Souhaite aider",
        "Besoin d’aide", "Personnes isolées", "Préférences communication",
        "Groupes locaux", "Personnes relais", "Avis projets",
        "Jeunes - Manques", "Seniors - Besoins", "Idée journée",
        "Téléphone", "Email", "Commentaire résident", "Commentaire prospecteur"
    ])
    wb.save(EXCEL_FILE)

@app.route('/', methods=['GET', 'POST'])
def questionnaire():
    if request.method == 'POST':
        data = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            request.form.get("nom"),
            request.form.get("prenom"),
            request.form.get("adresse"),
            request.form.get("ville"),
            request.form.get("code_postal"),
            request.form.get("rue"),
            request.form.get("anciennete"),
            request.form.get("plaisir"),
            request.form.get("ameliorer"),
            request.form.get("activite"),
            request.form.get("type_evenement"),
            request.form.get("aider"),
            request.form.get("besoin_aide"),
            request.form.get("isolées"),
            request.form.get("communication"),
            request.form.get("groupes"),
            request.form.get("relais"),
            request.form.get("avis"),
            request.form.get("manques"),
            request.form.get("seniors"),
            request.form.get("idee"),
            request.form.get("telephone"),
            request.form.get("email"),
            request.form.get("commentaire"),
            request.form.get("commentaire_prospecteur")
        ]
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append(data)
        wb.save(EXCEL_FILE)
        return redirect('/merci')
    return render_template('formulaire.html')

@app.route('/merci')
def merci():
    return "<h2 style='text-align:center'>✅ Merci, vos réponses ont bien été enregistrées !</h2>"

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
