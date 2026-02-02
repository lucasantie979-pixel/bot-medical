import telebot
import pandas as pd
from fpdf import FPDF
import os
from datetime import datetime

# --- CONFIGURATION ---
TOKEN = "8022132262:AAEJjnG344neYsB1RZk6qrz-54KDfjy3e6I"
bot = telebot.TeleBot(TOKEN)

# Nettoyage webhook
try: bot.remove_webhook()
except: pass

CONFIG = {
    "MEDECIN": "Dr Lahrach Ghizlane",
    "ADRESSE": "Avenue D, Secteur A, Imm 4, 1er étage, App n° 13, hay rahma salé",
    "PRIX": 80,
    "LOGO_LOCAL": "cnss.png" 
}

MOIS_FR = ['JANVIER','FEVRIER','MARS','AVRIL','MAI','JUIN','JUILLET','AOUT','SEPTEMBRE','OCTOBRE','NOVEMBRE','DECEMBRE']

class PDF_CNSS(FPDF):
    def header_custom(self, mois_ref):
        self.set_margins(10, 10, 10)
        self.set_auto_page_break(True, margin=15)
        
        # --- LARGEURS EXACTES (Total = 190mm) ---
        self.w_cols = [12, 63, 35, 35, 20, 25] 
        total_width = sum(self.w_cols)
        
        # --- EN-TÊTE ---
        self.set_xy(10, 10)
        h_box = 35 

        # 1. BLOC GAUCHE
        w_left = self.w_cols[0] + self.w_cols[1]
        self.set_line_width(0.4)
        self.rect(10, 10, w_left, h_box)
        
        if os.path.exists(CONFIG["LOGO_LOCAL"]):
            self.image(CONFIG["LOGO_LOCAL"], x=10 + (w_left - 25)/2, y=14, w=25)

        # 2. BLOC CENTRE
        x_center = 10 + w_left
        w_center = self.w_cols[2] + self.w_cols[3] + self.w_cols[4]
        self.rect(x_center, 10, w_center, h_box)

        self.set_xy(x_center, 12)
        self.set_font("Arial", 'B', 10)
        self.cell(w_center, 4, "ETAT DES FRAIS DE CONTROLES MEDICAUX", 0, 1, 'C')
        
        self.set_xy(x_center, 17)
        self.cell(w_center, 4, f"AU TITRE DU MOIS {mois_ref}", 0, 1, 'C')
        
        self.set_xy(x_center, 23)
        self.set_font("Arial", 'B', 7)
        self.cell(w_center, 3, "APPLICATION ART.32.37.47 ET 57 DU DAHIR", 0, 1, 'C')
        self.set_xy(x_center, 27)
        self.cell(w_center, 3, "PORTANT LOI N° 1.72.184 DU 15 JOUMADA II", 0, 1, 'C')
        self.set_xy(x_center, 31)
        self.cell(w_center, 3, "1392 (27 JUILLET 1972 )", 0, 1, 'C')

        # 3. BLOC DROITE
        x_right = x_center + w_center
        w_right = self.w_cols[5]
        self.rect(x_right, 10, w_right, h_box)
        
        self.set_xy(x_right, 12)
        self.set_font("Arial", 'B', 7)
        self.multi_cell(w_right, 3, "DIRECTION DES PRESTATIONS", 0, 'C')
        self.set_xy(x_right, 22)
        self.cell(w_right, 3, "---------------------", 0, 1, 'C')
        self.set_xy(x_right, 28)
        self.set_font("Arial", 'B', 8)
        self.cell(w_right, 3, "Réf : 312-2-05", 0, 1, 'C')

        # --- LIGNE MÉDECIN ---
        self.set_xy(10, 10 + h_box) 
        self.set_font("Arial", 'B', 8) 
        text_med = f"NOM DU MEDECIN CONTROLEUR : {CONFIG['MEDECIN']} / ADRESSE : {CONFIG['ADRESSE']}"
        self.cell(total_width, 8, text_med, 1, 1, 'C')

        # --- TITRES TABLEAU (CORRECTION ICI) ---
        # Police réduite à 7 pour que ça rentre bien
        self.set_font("Arial", 'B', 7)
        titles = ["N° ORDRE", "NOM PRENOM", "N° IMMATRICULATION", "N° DOSSIER", "MONTANT DES\nHONORAIRES", "NATURE DES\nPRESTATIONS (1)"]
        
        y_before = self.get_y()
        x_curr = 10
        h_title = 10 

        for i, title in enumerate(titles):
            self.set_xy(x_curr, y_before)
            # Dessine le cadre
            self.cell(self.w_cols[i], h_title, "", 1)
            
            # Positionnement précis du texte
            if '\n' in title:
                # Si 2 lignes : on commence un peu plus haut (2mm du bord) avec un interligne serré (3mm)
                self.set_xy(x_curr, y_before + 2)
                self.multi_cell(self.w_cols[i], 3, title, 0, 'C')
            else:
                # Si 1 ligne : on centre verticalement (3.5mm du bord)
                self.set_xy(x_curr, y_before + 3.5)
                self.multi_cell(self.w_cols[i], 3, title, 0, 'C')
            
            x_curr += self.w_cols[i]

        self.set_y(y_before + h_title)

# --- OUTILS ---
def trouver_mois(df, col_date):
    try:
        dates = pd.to_datetime(df[col_date], errors='coerce').dropna()
        if dates.empty: return "INCONNU"
        mode = dates.dt.to_period('M').mode()[0]
        return f"{MOIS_FR[mode.month - 1]} {mode.year}"
    except: return "INCONNU"

# --- BOT ---
@bot.message_handler(commands=['start'])
def welcome(m): 
    bot.reply_to(m, "👋 Bot CNSS V5 (Correction Titres).\nEnvoie ton fichier Excel.")

@bot.message_handler(content_types=['document'])
def handle_excel(message):
    files_to_clean = []
    try:
        file_name = message.document.file_name
        if not file_name.lower().endswith('.xlsx'):
            bot.reply_to(message, "⚠️ Fichier .xlsx uniquement.")
            return
        
        bot.reply_to(message, "⏳ Génération du PDF...")

        # Téléchargement
        file_info = bot.get_file(message.document.file_id)
        input_path = f"input_{file_name}"
        with open(input_path, 'wb') as f: f.write(bot.download_file(file_info.file_path))
        files_to_clean.append(input_path)

        # Lecture
        df = pd.read_excel(input_path)
        cols = {c.lower().strip(): c for c in df.columns}

        col_pres = next((v for k,v in cols.items() if "prestation" in k), None)
        col_prenom = next((v for k,v in cols.items() if ("pre" in k or "prénom" in k) and v != col_pres), None)
        col_nom = next((v for k,v in cols.items() if "nom" in k and "pre" not in k), None)
        col_date = next((v for k,v in cols.items() if "date" in k), None)
        col_imma = next((v for k,v in cols.items() if "imma" in k), None)
        col_dos = next((v for k,v in cols.items() if "dossier" in k), None)

        if not col_nom:
            bot.reply_to(message, "❌ Erreur: Colonne 'Nom' introuvable.")
            return

        # Génération
        pdf = PDF_CNSS()
        pdf.add_page()
        
        mois_ref = trouver_mois(df, col_date)
        pdf.header_custom(mois_ref)

        pdf.set_font("Arial", '', 8)
        w = pdf.w_cols
        count = 1
        total = 0

        for _, row in df.iterrows():
            nom = str(row[col_nom])
            if pd.isna(row[col_nom]) or len(nom.strip()) < 2 or str(row[col_nom]).lower() == 'nan': continue
            
            prenom = str(row[col_prenom]) if col_prenom and not pd.isna(row[col_prenom]) else ""
            full_name = f"{nom} {prenom}".strip().upper()
            imma = str(row[col_imma]) if col_imma and not pd.isna(row[col_imma]) else ""
            dos = str(row[col_dos]) if col_dos and not pd.isna(row[col_dos]) else ""
            pres = str(row[col_pres]) if col_pres and not pd.isna(row[col_pres]) else ""

            h = 6
            pdf.cell(w[0], h, str(count), 1, 0, 'C')
            pdf.cell(w[1], h, full_name[:35], 1, 0, 'L') 
            pdf.cell(w[2], h, imma, 1, 0, 'C')
            pdf.cell(w[3], h, dos, 1, 0, 'C')
            pdf.cell(w[4], h, f"{CONFIG['PRIX']:.2f}", 1, 0, 'C')
            pdf.cell(w[5], h, pres[:15], 1, 1, 'C')

            total += CONFIG['PRIX']
            count += 1

        # Total
        pdf.set_font("Arial", 'B', 10)
        w_tot = sum(w[:4])
        pdf.cell(w_tot, 8, "TOTAL", 1, 0, 'C')
        pdf.cell(w[4], 8, f"{total:.2f}", 1, 0, 'C')
        pdf.set_font("Arial", '', 8)
        pdf.cell(w[5], 8, "DH", 1, 1, 'L')

        # Envoi
        output = f"Frais_{file_name.replace('.xlsx','.pdf')}"
        pdf.output(output)
        files_to_clean.append(output)

        with open(output, 'rb') as f: bot.send_document(message.chat.id, f)
        print(f"✅ PDF envoyé : {file_name}")

    except Exception as e:
        print(e)
        bot.reply_to(message, f"❌ Erreur: {e}")
    finally:
        for f in files_to_clean: 
            if os.path.exists(f): os.remove(f)

print("✅ Bot CNSS en ligne (V5)...")
bot.infinity_polling()
