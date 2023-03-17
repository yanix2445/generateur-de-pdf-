import os
import tkinter as tk
from tkinter import filedialog
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# Fonction pour afficher une boîte de dialogue permettant à l'utilisateur de choisir un emplacement pour enregistrer le fichier PDF


def get_save_location():
    root = tk.Tk()
    root.withdraw()
    save_location = filedialog.asksaveasfilename(
        defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    return save_location


def create_invoice(output_file, invoice_data):
    # Créez un objet SimpleDocTemplate avec les marges et la taille de la page spécifiées
    doc = SimpleDocTemplate(output_file, pagesize=letter,
                            rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    elements = []

    # Styles
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Center', alignment=TA_CENTER))
    styles.add(ParagraphStyle(name='Left', alignment=TA_LEFT))

    # Ajouter le logo de l'entreprise (si le fichier 'mon_logo.png' existe)
    logo_path = r'C:\Users\Username\Documents\logo.png'
    if os.path.isfile(logo_path):
        logo = logo_path
        elements.append(logo)

    # Titre de la facture
    elements.append(Paragraph("Facture", styles["Heading1"]))
    elements.append(Spacer(1, 24))

    # Informations sur l'entreprise
    elements.append(Paragraph("Nom de l'entreprise", styles["Normal"]))
    elements.append(Paragraph("Adresse", styles["Normal"]))
    elements.append(Paragraph("Ville, Code postal", styles["Normal"]))
    elements.append(Paragraph("Téléphone: 0123456789", styles["Normal"]))
    elements.append(Spacer(1, 24))

    # Informations sur le client
    elements.append(Paragraph(f"Facturé à :", styles["Normal"]))
    elements.append(
        Paragraph(f"{invoice_data['client_name']}", styles["Normal"]))
    elements.append(
        Paragraph(f"{invoice_data['client_address']}", styles["Normal"]))
    elements.append(Spacer(1, 24))

    # Informations sur la facture
    elements.append(Paragraph(
        f"Numéro de facture : {invoice_data['invoice_number']}", styles["Normal"]))
    elements.append(Paragraph(
        f"Date de la facture : {invoice_data['invoice_date']}", styles["Normal"]))
    elements.append(Spacer(1, 24))

    # Créer un tableau pour les articles de la facture
    table_data = [['Description', 'Quantité', 'Prix unitaire', 'Total']]

    # Ajouter les articles de la facture au tableau
    for item in invoice_data['items']:
        table_data.append(
            [item['description'], item['quantity'], item['unit_price'], item['total']])

    # Créer un objet Table à partir des données du tableau et appliquer un style
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 12),
        ('TOPPADDING', (0, 1), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 12),
    ]))

    # Ajouter le tableau à la liste des éléments de la facture
    elements.append(table)
    elements.append(Spacer(1, 12))

    # Ajouter le total HT, la TVA et le total TTC à la facture
    elements.append(
        Paragraph(f"Total (HT) : {invoice_data['total_ht']} €", styles["Normal"]))
    elements.append(Paragraph(
        f"TVA ({invoice_data['tva_rate']}%) : {invoice_data['tva']} €", styles["Normal"]))
    elements.append(
        Paragraph(f"Total (TTC) : {invoice_data['total_ttc']} €", styles["Normal"]))

    # Générer la facture en PDF en ajoutant les éléments à l'objet SimpleDocTemplate
    doc.build(elements)


if __name__ == "__main__":
    # Exemple de données pour la facture
    invoice_data = {
        "invoice_number": "INV-001",
        "invoice_date": "01/01/2023",
        "client_name": "John Doe",
        "client_address": "123 Main St, City, ZIP",
        "items": [
            {"description": "Produit 1", "quantity": 2,
                "unit_price": 50, "total": 100},
            {"description": "Produit 2", "quantity": 1,
                "unit_price": 75, "total": 75},
        ],
        "total_ht": 175,
        "tva_rate": 20,
        "tva": 35,
        "total_ttc": 210,
    }

    # Obtenir l'emplacement de sauvegarde choisi par l'utilisateur
    save_location = get_save_location()

    # Vérifier si l'utilisateur a sélectionné un emplacement de sauvegarde
    if save_location:
        # Créer la facture en PDF à l'emplacement choisi avec les données fournies
        create_invoice(save_location, invoice_data)

        # Afficher un message indiquant que la facture a été créée avec succès
        print(f"La facture a été créée avec succès : {save_location}")
    else:
        # Afficher un message indiquant que l'utilisateur n'a pas sélectionné d'emplacement de sauvegarde et que l'opération a été annulée
        print("Aucun emplacement de sauvegarde sélectionné. Opération annulée.")
