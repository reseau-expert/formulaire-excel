import requests
from flask import Flask, render_template, request, redirect
from openpyxl import Workbook, load_workbook
import os
from datetime import datetime

ZAPIER_WEBHOOK_URL = "https://hooks.zapier.com/hooks/catch/23360947/uy09xop/"
EXCEL_FILE = "reponses.xlsx"

app = Flask(__name__)

# Création du fichier Excel s’il n’existe pas
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.append([
        "Date", "Nom", "Prénom", "Adresse", "Ville", "Code postal", "Rue ou résidence", "Ancienneté", "Ce que j’aime", "À améliorer",
        "Activité souhaitée", "Type d’événements", "Souhaite aider", "Personnes aidées", "Personnes isolées", "Préférences communication",
        "Groupes locaux", "Personnes relais", "Avis projets", "Besoins manques", "Seniors - Besoins", "Idée journée",
        "Téléphone", "Email", "Commentaire résident", "Commentaire prospecteur"
    ])
    wb.save(EXCEL_FILE)

@app.route("/", methods=["GET", "POST"])
def questionnaire():
    if request.method == "POST":
        data = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            request.form.get("nom"),
            request.form.get("prenom"),
            request.form.get("adresse"),
            request.form.get("ville"),
            request.form.get("code_postal"),
            request.form.get("rue"),
            request.form.get("ancienne"),
            request.form.get("plaisir"),
            request.form.get("ameliorer"),
            request.form.get("activite"),
            request.form.get("type_evenement"),
            request.form.get("aider"),
            request.form.get("besoin_aide"),
            request.form.get("isolees"),
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

        # Enregistrement dans Excel
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append(data)
        wb.save(EXCEL_FILE)

        # Envoi vers Zapier
        try:
            requests.post(ZAPIER_WEBHOOK_URL, json={
                "nom": request.form.get("nom"),
                "prenom": request.form.get("prenom"),
                "ville": request.form.get("ville"),
                "code_postal": request.form.get("code_postal"),
                "telephone": request.form.get("telephone"),
                "email": request.form.get("email"),
                "commentaire": request.form.get("commentaire"),
                "commentaire_prospecteur": request.form.get("commentaire_prospecteur")
            })
        except Exception as e:
            print("Erreur Zapier:", e)

        return redirect("/merci")

    return render_template("formulaire.html")

@app.route("/merci")
def merci():
    return "<h2 style='text-align:center;'>✅ Merci, vos réponses ont bien été enregistrées </h2>"

if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)

