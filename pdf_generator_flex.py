import os
import random
import string
import tkinter as tk
from tkinter import filedialog, simpledialog
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship, sessionmaker
from sqlalchemy import Column, Integer, String, ForeignKey
from sqlalchemy.orm import declarative_base, relationship

# Modèles SQLAlchemy
Base = declarative_base()


def create_invoice_ui():
    root = tk.Tk()
    root.withdraw()

    client_id = simpledialog.askinteger("Client ID", "Entrez l'ID du client:")

    if client_id:
        invoice_data = get_invoice_data(client_id)

        if invoice_data:
            save_location = get_save_location()

            if save_location:
                create_invoice(save_location, invoice_data)
                print(f"La facture a été créée avec succès : {save_location}")
            else:
                print("Aucun emplacement de sauvegarde sélectionné. Opération annulée.")
        else:
            print(
                f"Aucune facture trouvée pour le client avec l'ID {client_id}")
    else:
        print("Aucun ID de client entré. Opération annulée.")


def create_client():
    def submit_new_client():
        name = name_entry.get().strip()
        address = address_entry.get().strip()

        if name and address:
            session = Session()
            new_client = Client(name=name, address=address)
            session.add(new_client)
            session.commit()
            session.close()

            print(f"Client créé avec succès : {name}")
            new_client_window.destroy()
        else:
            print("Veuillez entrer un nom et une adresse valides.")

    new_client_window = tk.Toplevel()
    new_client_window.title("Créer un nouveau client")

    name_label = tk.Label(new_client_window, text="Nom:")
    name_label.grid(row=0, column=0, padx=(10, 0), pady=(10, 5))
    name_entry = tk.Entry(new_client_window)
    name_entry.grid(row=0, column=1, padx=(0, 10), pady=(10, 5))

    address_label = tk.Label(new_client_window, text="Adresse:")
    address_label.grid(row=1, column=0, padx=(10, 0), pady=(0, 5))
    address_entry = tk.Entry(new_client_window)
    address_entry.grid(row=1, column=1, padx=(0, 10), pady=(0, 5))

    submit_button = tk.Button(
        new_client_window, text="Créer", command=submit_new_client)
    submit_button.grid(row=2, columnspan=2, pady=(10, 10))


def create_new_client():
    # Récupérer le dernier ID de client utilisé
    session = Session()
    last_client = session.query(Client).order_by(Client.id.desc()).first()
    new_client_id = last_client.id + 1 if last_client else 1

    # Créer un nouveau client avec le nouvel ID de client
    new_client_name = simpledialog.askstring("Nouveau client", "Entrez le nom du client:")
    new_client_address = simpledialog.askstring("Nouveau client", "Entrez l'adresse du client:")

    if new_client_name and new_client_address:
        new_client = Client(id=new_client_id, name=new_client_name, address=new_client_address)
        new_invoice = Invoice(id=new_client_id, invoice_number='INV-'+str(new_client_id), amount=0, client_id=new_client_id)
        session.add(new_client)
        session.add(new_invoice)
        session.commit()
        print(f"Le nouveau client a été créé avec succès. ID du client : {new_client_id}")
    else:
        print("Aucune information sur le client entrée. Opération annulée.")


def main():
    # Créez l'interface utilisateur pour créer de nouveaux clients
    root = tk.Tk()
    root.title("Générateur de factures")

    create_new_client_button = tk.Button(
        root, text="Créer un nouveau client", command=create_new_client)
    create_new_client_button.pack(pady=10)

    create_invoice_button = tk.Button(
        root, text="Créer une facture", command=create_invoice_ui)
    create_invoice_button.pack(pady=10)

    root.mainloop()


class Client(Base):
    __tablename__ = 'clients'
    id = Column(Integer, primary_key=True)
    name = Column(String)
    address = Column(String)
    invoices = relationship("Invoice", back_populates="client")

    def __repr__(self):
        return f"Client(id={self.id}, name='{self.name}', address='{self.address}')"

class Invoice(Base):
    __tablename__ = 'invoices'
    id = Column(Integer, primary_key=True)
    invoice_number = Column(String)
    amount = Column(Integer)
    client_id = Column(Integer, ForeignKey('clients.id'))  # Modifié ici
    client = relationship("Client", back_populates="invoices")

    def __repr__(self):
        return f"Invoice(id={self.id}, invoice_number='{self.invoice_number}', amount={self.amount}, client_id={self.client_id})"




# Créer une connexion à la base de données SQLite et créer les tables
engine = create_engine('sqlite:///invoices.db')
Base.metadata.create_all(engine)

# Créer une session pour interagir avec la base de données
Session = sessionmaker(bind=engine)


def get_invoice_data(client_id):
    session = Session()
    invoice = session.query(Invoice).filter_by(id=client_id).first()

    if invoice:
        # Récupérer les données de la facture pour le client sélectionné
        # À adapter en fonction de la structure de votre base de données
        invoice_data = {
            'invoice_number': invoice.id,
            'invoice_date': '01/01/2023',  # À adapter
            'client_name': invoice.client.name,
            'client_address': invoice.client.address,
            'items': [
                {"description": "Produit 1", "quantity": 2,
                    "unit_price": 50, "total": 100},
                {"description": "Produit 2", "quantity": 1,
                    "unit_price": 75, "total": 75},
            ],
            'total_ht': 175,
            'tva_rate': 20,
            'tva': 35,
            'total_ttc': 210
        }
        return invoice_data
    else:
        return None




def get_save_location():
    root = tk.Tk()
    root.withdraw()
    save_location = filedialog.asksaveasfilename(
        defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    return save_location


def create_invoice(output_file, invoice_data):
    doc = SimpleDocTemplate(output_file, pagesize=letter)
    styles = getSampleStyleSheet()
    data = []

    # Titre de la facture
    title_style = ParagraphStyle(
        'title', parent=styles['Title'], fontSize=24, alignment=TA_CENTER)
    title = Paragraph("Facture", title_style)
    data.append([title])

    # Informations sur la facture et le client
    client_info = Paragraph(
        f"Numéro de facture : {invoice_data['invoice_number']}<br/>"
        f"Date de facture : {invoice_data['invoice_date']}<br/>"
        f"<br/>"
        f"Nom du client : {invoice_data['client_name']}<br/>"
        f"Adresse du client : {invoice_data['client_address']}",
        styles["Normal"]
    )
    data.append([client_info])

    # Ajouter un espace
    spacer = Spacer(1, 20)
    data.append([spacer])

    # Détails des articles
    items_header = ["Description", "Quantité", "Prix unitaire", "Total"]
    items_data = [items_header]
    for item in invoice_data["items"]:
        items_data.append(
            [item["description"], item["quantity"],
                item["unit_price"], item["total"]]
        )
    items_table = Table(items_data)
    items_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
            ]
        )
    )
    data.append([items_table])

    # Ajouter un espace
    data.append([spacer])

    # Récapitulatif des montants
    recap = Paragraph(
        f"Total HT : {invoice_data['total_ht']} €<br/>"
        f"TVA ({invoice_data['tva_rate']} %) : {invoice_data['tva']} €<br/>"
        f"Total TTC : {invoice_data['total_ttc']} €",
        styles["Normal"],
    )
    data.append([recap])

    # Créer un tableau pour organiser les éléments de la facture
    table = Table(data, colWidths=[doc.width])
    table.setStyle(
        TableStyle(
            [
                ("TOPPADDING", (0, 0), (-1, -1), 12),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
            ]
        )
    )

    # Ajouter la table au document et construire le document
    doc.build([table])


if __name__ == "__main__":
    main()
