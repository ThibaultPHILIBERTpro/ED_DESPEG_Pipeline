"""
=============================================================================
PROJET : Pipeline ETL - Analyse et Suivi des Doctorants (ED DESPEG)
AUTEUR : PHILIBERT Thibault
DATE DE MISE À JOUR : Mai 2026

NOTE DE PORTEFEUILLE (PROPRIÉTÉ INTELLECTUELLE & RGPD) :
Ce fichier présente une sélection de fonctions authentiques extraites du 
moteur ETL (> 1 200 lignes) développé pour l'École Doctorale DESPEG.
L'orchestration globale, les dictionnaires de contrôle interne et les 
chemins d'extraction ADUM ont été retirés par respect pour la confidentialité 
de l'établissement.

Ces extraits illustrent 5 briques techniques :
1. Sécurité Système : Airbag anti-crash (sys.excepthook) & Vigile OS (PermissionError).
2. Interface Graphique (CustomTkinter) : Menu principal et centrage dynamique.
3. Moteur ETL & Regex : Nettoyage Unicode (NFKD), dictionnaire de synonymes et arbre N4.
4. Automatisation Excel Avancée (XlsxWriter) : Colonnes fantômes, menus déroulants
en cascade (INDIRECT/SUBSTITUTE) et mise en forme conditionnelle (Jaune/Rouge).
5. Modélisation BI : Calcul vectorisé des 14 KPIs et génération de clé composite.
=============================================================================
"""

import os                                          # Pour vérifier l'existence des dossiers sur l'ordinateur
import sys                                         # Permet de forcer l'arrêt total du script si l'utilisateur annule
import traceback
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk                        # Pour créer toutes les interfaces graphiques
import pandas as pd                                # Pour lire, modifier et exporter les fichiers Excel
import numpy as np                                 # Pour les conditions complexes
import datetime                                    # Pour horodater les fichiers de sauvegarde d'erreurs
import shutil                                      # Outil pour copier-coller les fichiers
from xlsxwriter.utility import xl_col_to_name      # Outil pour convertir les index
import unicodedata                                 # Outil pour gérer les accents


# ===================================================================
# 1) SÉCURITÉ SYSTÈME : AIRBAG D'URGENCE & VIGILE FICHIERS VERROUILLÉS
# ===================================================================

def airbag_urgence(exc_type, exc_value, exc_traceback):
    """
    S'active UNIQUEMENT si le script fait un crash total non prévu.
    """
    # Si le script s'arrête normalement ou via un de tes sys.exit(0), on ne fait rien
    if issubclass(exc_type, SystemExit):
        return

    # On récupère le texte rouge classique de l'erreur Python (le fameux code d'erreur)
    erreur_texte = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    
    # On essaie d'écrire un fichier texte sur son bureau ou dans le dossier problemes
    try:
        os.makedirs("data/problemes", exist_ok=True)
        horodatage = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        chemin_log = f"data/problemes/CRASH_IMPREVU_{horodatage}.txt"
        with open(chemin_log, "w", encoding="utf-8") as f:
            f.write("=== CRASH IMPREVU ===\n")
            f.write(erreur_texte)
    except:
        chemin_log = "impossible d'écrire le log."

    # On lance la fenêtre d'alerte Windows de base (Tkinter)
    root = tk.Tk()
    root.withdraw() # Cache la fenêtre de fond
    root.attributes('-topmost', True) # Force au premier plan
    
    message_alerte = (
        "Le pipeline a rencontré une erreur imprévue et s'est arrêté.\n\n"
        f"Un rapport a été généré ici : {chemin_log}\n\n"
        "Message technique (à envoyer à Thibault) :\n" + str(exc_value)
    )
    
    messagebox.showerror("CRASH CRITIQUE", message_alerte)
    root.destroy()

# On connecte l'airbag au moteur Python
sys.excepthook = airbag_urgence


def verifier_fichiers_fermes(fichiers_cibles):
    """
    Vérifie si les fichiers Excel cibles sont ouverts par un autre programme.
    Boucle à l'infini tant que l'utilisateur n'a pas fermé les fichiers.
    """
    while True:
        fichiers_bloques = []
        for f in fichiers_cibles:
            if os.path.exists(f):
                try:
                    # Tente d'ouvrir en mode ajout pour tester le verrou Windows
                    with open(f, 'a'):
                        pass
                except PermissionError:
                    fichiers_bloques.append(os.path.basename(f))
        
        # On enlève les doublons potentiels
        fichiers_bloques = list(set(fichiers_bloques))
        
        if not fichiers_bloques:
            break # Tout est bon, on sort de la boucle !
        
        # Interface d'alerte bloquante
        fenetre_lock = ctk.CTk()
        fenetre_lock.title("Erreur : Fichier(s) ouvert(s)")
        
        w_lock, h_lock = 600, 300
        x_lock = int((fenetre_lock.winfo_screenwidth() / 2) - (w_lock / 2))
        y_lock = int((fenetre_lock.winfo_screenheight() / 2) - (h_lock / 2))
        fenetre_lock.geometry(f"{w_lock}x{h_lock}+{x_lock}+{y_lock}")
        fenetre_lock.attributes('-topmost', True)
        
        def reessayer():
            fenetre_lock.destroy() # Ça va relancer le While True
            
        def annuler():
            fenetre_lock.destroy()
            sys.exit(0) # Arrêt total
        
        fenetre_lock.protocol("WM_DELETE_WINDOW", annuler)
        
        ctk.CTkLabel(fenetre_lock, text="Fichiers Excel verrouillés", font=("Roboto", 16, "bold"), text_color="red").pack(pady=(20, 10))
        ctk.CTkLabel(fenetre_lock, text="Veuillez fermer les fichiers suivants pour que le script puisse les mettre à jour :\n\n" + "\n".join(f"- {f}" for f in fichiers_bloques), font=("Roboto", 13)).pack(pady=10)
        
        frame_btn = ctk.CTkFrame(fenetre_lock, fg_color="transparent")
        frame_btn.pack(pady=20)
        ctk.CTkButton(frame_btn, text="Quitter le script", fg_color="gray", command=annuler).pack(side="left", padx=10)
        ctk.CTkButton(frame_btn, text="J'ai fermé, Réessayer", command=reessayer).pack(side="right", padx=10)
        
        fenetre_lock.mainloop()


# ===================================================================
# 2) INTERFACE GRAPHIQUE : MENU PRINCIPAL (CUSTOMTKINTER)
# ===================================================================

def afficher_menu_principal():
    """
    Ce menu permet de choisir quel mode d'exécution on veut lancer :
    - MODULE STATS : Focus sur les bases 1 à 4 (Inscrits, Abandons, Soutenances).
    - MODULE ALUMNI : Focus sur l'insertion pro (T+1, T+3, T+5) et Limesurvey.
    """
    resultat = {"mode": None}
    fen = ctk.CTk()
    fen.title("Menu Principal - Pipeline Data DESPEG")
    fen.geometry("650x650") # Fenêtre agrandie pour les 5 boutons
    
    # Centrage propre
    largeur_ecran, hauteur_ecran = fen.winfo_screenwidth(), fen.winfo_screenheight()
    fen.geometry(f"+{int((largeur_ecran/2)-(650/2))}+{int((hauteur_ecran/2)-(650/2))}")
    fen.focus_force()
    
    # --- Définition des 5 actions ---
    def choix_stats_1():
        resultat["mode"] = "STATS_BROUILLON"
        fen.destroy()
    def choix_stats_2():
        resultat["mode"] = "STATS_VALIDATION"
        fen.destroy()
    def choix_alumni_1():
        resultat["mode"] = "ALUMNI_BROUILLON"
        fen.destroy()
    def choix_alumni_relance():
        resultat["mode"] = "ALUMNI_RELANCE"
        fen.destroy()
    def choix_alumni_2():
        resultat["mode"] = "ALUMNI_VALIDATION"
        fen.destroy()
    def annuler():
        fen.destroy()
        sys.exit(0)
        
    fen.protocol("WM_DELETE_WINDOW", annuler)
    ctk.CTkLabel(fen, text="Sélectionnez le mode d'exécution :", font=("Roboto", 18, "bold")).pack(pady=(20, 10))
    
    # ==========================================
    # BLOC 1 : STATISTIQUES (Mars/Avril)
    # ==========================================
    ctk.CTkLabel(fen, text="MODULE STATISTIQUES (Mars/Avril)", font=("Roboto", 14, "bold"), text_color="#1f538d").pack(pady=(10, 5))
    
    ctk.CTkButton(fen, text="1. Créer les bases Stats", command=choix_stats_1, width=350, height=40).pack(pady=5)
    ctk.CTkLabel(fen, text="Extrait les informations des bases ADUM, nettoie et génère les Excel à corriger.", text_color="gray", font=("Roboto", 11)).pack()
    
    ctk.CTkButton(fen, text="2. Valider et Archiver Stats", command=choix_stats_2, width=350, height=40, fg_color="#27AE60", hover_color="#1E8449").pack(pady=(10, 5))
    ctk.CTkLabel(fen, text="Fige les indicateurs et archive la rentrée.", text_color="gray", font=("Roboto", 11)).pack()
    
    # --- Ligne de séparation visuelle ---
    ctk.CTkFrame(fen, width=400, height=2, fg_color="gray").pack(pady=15)
    
    # ==========================================
    # BLOC 2 : ALUMNI & INSERTION PRO (Juin)
    # ==========================================
    ctk.CTkLabel(fen, text="MODULE INSERTION PRO (Juin)", font=("Roboto", 14, "bold"), text_color="#E67E22").pack(pady=(5, 5))
    
    ctk.CTkButton(fen, text="3. Préparer l'Enquête Alumni", command=choix_alumni_1, width=350, height=40).pack(pady=5)
    ctk.CTkLabel(fen, text="Génère les tableaux de saisie (T+1, T+3, T+5) et les CSV Limesurvey.", text_color="gray", font=("Roboto", 11)).pack()
    
    # NOUVEAU BOUTON RELANCE
    ctk.CTkButton(fen, text="4. Générer les relances Directeurs", command=choix_alumni_relance, width=350, height=40, fg_color="#F39C12", hover_color="#D68910").pack(pady=5)
    ctk.CTkLabel(fen, text="Extrait les e-mails des directeurs pour les anciens doctorants injoignables.", text_color="gray", font=("Roboto", 11)).pack()

    ctk.CTkButton(fen, text="5. Intégrer et Archiver Alumni", command=choix_alumni_2, width=350, height=40, fg_color="#27AE60", hover_color="#1E8449").pack(pady=(10, 5))
    ctk.CTkLabel(fen, text="Met à jour la mémoire et nettoie les fichiers temporaires.", text_color="gray", font=("Roboto", 11)).pack()
    
    fen.mainloop()
    return resultat["mode"]


# ===================================================================
# 3) FONCTIONS UTILITAIRES & MOTEUR DE CLASSIFICATION REGEX (N1-N4)
# ===================================================================

pays_ue_27 = [
    "allemagne", "allemand", "allemande", "autriche", "autrichien", "autrichienne",
    "belgique", "belge", "bulgarie", "bulgare", "chypre", "chypriote",
    "croatie", "croate", "danemark", "danois", "danoise", "espagne", "espagnol", "espagnole",
    "estonie", "estonien", "estonienne", "finlande", "finlandais", "finlandaise",
    "grece", "grec", "grecque", "hongrie", "hongrois", "hongroise",
    "irlande", "irlandais", "irlandaise", "italie", "italien", "italienne",
    "lettonie", "letton", "lettonne", "lituanie", "lituanien", "lituanienne",
    "luxembourg", "luxembourgeois", "luxembourgeoise", "malte", "maltais", "maltaise",
    "pays-bas", "neerlandais", "neerlandaise", "hollandais", "hollandaise",
    "pologne", "polonais", "polonaise", "portugal", "portugais", "portugaise",
    "republique tcheque", "tcheque", "roumanie", "roumain", "roumaine",
    "slovaquie", "slovaque", "slovenie", "slovene", "suede", "suedois", "suedoise"
]

def lisser_texte(colonne_pandas, valeur_na=''):
    """
    Utilisée pour standardiser les textes afin de retirer les majuscules, les accents et les espaces inutiles.
    Remplace également les valeurs manquantes par une chaîne vide (par défaut) ou par toute autre valeur si spécifiée lors de l'appel de la fonction.
    """
    return (
        colonne_pandas
        .fillna(valeur_na)
        .astype(str)
        .str.lower()
        .str.normalize('NFKD')
        .str.encode('ascii', errors='ignore')
        .str.decode('utf-8')
        .str.strip()
    )

def standardiser_colonnes(df):
    """
    Dictionnaire des synonymes administratif : ADUM modifie parfois le nom de ses 
    colonnes d'une année sur l'autre ou selon l'extraction. Cette fonction force 
    les colonnes clés à adopter notre nomenclature standard pour éviter que le script ne casse.
    """
    if df is None or df.empty: return df
    mapping = {}
    for col in df.columns:
        col_clean = lisser_texte(pd.Series([col])).iloc[0]
        if col_clean in ["n adum", "no adum", "numero adum", "adum", "matricule", "matricule etude", "matricule etudiant"]:
            mapping[col] = "Matricule"
        elif col_clean in ["nom de naissance", "nom d'usage", "nom de famille", "nom famille", "nom"]:
            mapping[col] = "Nom"
        elif col_clean in ["prenom", "prénom"]:
            mapping[col] = "Prenom"
    return df.rename(columns=mapping)

def securiser_matricule(df):
    """
    Nettoie la variable Matricule pour éviter les bugs de fusion (espaces, .0 à la fin).
    Retourne le dataframe intact si le fichier source était manquant ou vide.
    """
    if df is None or df.empty: 
        return df
        
    if 'Matricule' in df.columns:
        df['Matricule'] = df['Matricule'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    return df

def appliquer_moteur_n4(colonne_texte, colonne_nat, colonne_employeur):
    """
    Moteur logique classifiant les sources de financement en catégories N4 strictes.
    Utilise des expressions régulières (Regex) pour scanner les textes bruts d'ADUM
    et croise l'information avec la nationalité et le type d'employeur.
    """
    texte_cible = lisser_texte(colonne_texte, valeur_na='')
    nat_propre = lisser_texte(colonne_nat, valeur_na='')
    employeur_propre = lisser_texte(colonne_employeur, valeur_na='')
    
    est_ue = nat_propre.isin(pays_ue_27) | nat_propre.isin(['france', 'francais', 'francaise'])
    regex_inter = r'\berc\b|marie sklodowska-curie|erasmus|union europeenne|hors pays membres|etranger|organisme international|campus france|organismes internationaux'
    
    regex_pro = r'\bcadre\b|emploi a titre principal|activite remuneree autre|fonctionnaire|hors financement these'
    regex_non_fi = r'sans activite|sans financement|non concerne|aucun|non finance'
    
    # On repère les vides et les "non concerné" dans la colone employeur
    employeur_absent = employeur_propre.isin(['', 'nan', 'none', 'nd',"-"]) | employeur_propre.str.contains('non concerne', regex=True)
    # On regroupe tout ce qui indique un contrat privé
    regex_prive = r"skema|edhec|ecole privee|contrats doctoraux de droit prive|contrat doctoral de droit prive|partenariat|\bgip\b|mecenat"
    masque_prive = texte_cible.str.contains(regex_prive, regex=True)
    
    # 2. On cherche les mots-clés d'écoles dans l'employeur (ou dans le texte de base au cas où)
    regex_ecole = r"skema|edhec|ecole|business school|\bbs\b|universit|univ|institut"
    masque_ecole = employeur_propre.str.contains(regex_ecole, regex=True) | texte_cible.str.contains(regex_ecole, regex=True)
    
    conds = [
        (texte_cible.str.contains(r'\bcifre\b', regex=True)),  
        (texte_cible.str.contains(r'\bidex\b|\bisite\b|\bpia\b|france 2030|labex', regex=True)), 
        (texte_cible.str.contains(r'\banr\b', regex=True)), 
        (texte_cible.str.contains(r"\bmesr\b|\bepscp\b|\bepst\b|\bepic\b|\bepa\b|collectivite|region|\bcofra\b|financement d'un etablissement public francais|programme specifique handicap|\bmesr\b|ministere de l'enseignement superieur et de la recherche|\bmesri\b|programmes specifiques normaliens et polytechniciens|programmes ministeriels specifiques", regex=True)), 
        
        (masque_prive & masque_ecole),   # C'est privé et c'est une école
        (masque_prive & ~masque_ecole),  # C'est privé mais c'est pas une école
        
        (texte_cible.str.contains(regex_inter, regex=True) & est_ue),  
        (texte_cible.str.contains(regex_inter, regex=True) & ~est_ue), 
        (texte_cible.str.contains(regex_pro, regex=True) & ~employeur_absent),
        (texte_cible.str.contains(regex_non_fi, regex=True) | (texte_cible.str.contains(regex_pro, regex=True) & employeur_absent))
    ]
    
    # On a supprimé "Autre privé" de la liste vu qu'on l'a fusionné avec "Autre partenariat"
    choix = [
        "CIFRE", "Idex/PIA", "ANR", "Ministère", 
        "Ecoles privées", "Autre partenariat",
        "UE (27)", "Hors UE", "Professionnel", "Non financé"
    ]
    return np.select(conds, choix, default="A Vérifier")


# ===================================================================
# 4) EXPORTATION INTERACTIVE AVEC EXCEL (COLONNES FANTÔMES & CASCADE)
# ===================================================================

def exporter_base_actifs_interactive(df, chemin_fichier, nom_feuille="Actifs"):
    """
    Création des colonnes fantômes pour l'audit de cohérence. Ces colonnes seront masquées dans Excel.
    Elles servent de point de comparaison pour la mise en forme conditionnelle
    La fonction permet aussi de bien structurer le fichier Excel avec un onglet caché pour les listes de validation et de créer les menus déroulants en cascade.
    Cette fonction est utiliser pour la base de données 1
    """
    nb_colonnes_visibles = len(df.columns)
    
    df['Ghost_N1'] = df['Financement_N1']
    df['Ghost_N2'] = df['Financement_N2']
    df['Ghost_N3'] = df['Financement_N3']
    df['Ghost_N4'] = df['Financement_N4']
    
    nb_colonnes_totales = len(df.columns)

    # 2) Création du fichier Excel
    writer = pd.ExcelWriter(chemin_fichier, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=nom_feuille, index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets[nom_feuille]

    # Mise sous forme de tableau (Design)
    (max_row, max_col) = df.shape
    column_settings = [{'header': column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {
        'columns': column_settings,
        'style': 'Table Style Medium 1'
    })
    worksheet.autofit()

    # On masque les colonnes fantômes (Ni vu ni connu)
    worksheet.set_column(nb_colonnes_visibles, nb_colonnes_totales - 1, None, None, {'hidden': 1})

    # 3) ONGLET CACHÉ : Les listes de référence
    worksheet_ref = workbook.add_worksheet('Ref_Listes')
    worksheet_ref.hide() # On cache l'onglet à l'utilisateur

    # Dictionnaire de toute ton architecture financière (Espaces remplacés par des tirets du bas)
    dict_listes = {
        "N1_Liste": ["Financé", "Non financé", "A Vérifier"],
        
        "Financé": ["Contrat doctoral", "Financement hors contrat doctoral"],
        "Non_financé": ["Non financé"],
        "A_Vérifier": ["A Vérifier"],
        
        "Contrat_doctoral": ["Contrat doctoral de droit public", "Contrat doctoral de droit privé", "CIFRE"],
        "Financement_hors_contrat_doctoral": ["Financements internationaux", "Professionnel"],
        
        "Contrat_doctoral_de_droit_public": ["Idex/PIA", "ANR", "Ministère"],
        "Contrat_doctoral_de_droit_privé": ["Ecoles privées", "Autre partenariat"],
        "CIFRE": ["CIFRE"],
        "Financements_internationaux": ["UE (27)","Hors UE"],
        "Professionnel": ["Professionnel"]
    }

    # On écrit ces listes dans l'onglet caché et on crée les "Noms" Excel
    col_idx = 0
    for nom_liste, valeurs in dict_listes.items():
        worksheet_ref.write_string(0, col_idx, nom_liste) # L'en-tête
        worksheet_ref.write_column(1, col_idx, valeurs)   # Les éléments
        
        lettre_col = chr(65 + col_idx) # Convertit 0 en A, 1 en B, etc.
        plage = f"=Ref_Listes!${lettre_col}$2:${lettre_col}${len(valeurs)+1}"
        workbook.define_name(nom_liste, plage) # Création du Named Range
        col_idx += 1

    # 4) Où sont les colonnes dans l'Excel ?
    idx_n1 = df.columns.get_loc('Financement_N1')
    idx_n2 = df.columns.get_loc('Financement_N2')
    idx_n3 = df.columns.get_loc('Financement_N3')
    idx_n4 = df.columns.get_loc('Financement_N4')

    idx_g1 = df.columns.get_loc('Ghost_N1')
    idx_g2 = df.columns.get_loc('Ghost_N2')
    idx_g3 = df.columns.get_loc('Ghost_N3')
    idx_g4 = df.columns.get_loc('Ghost_N4')

    # Convertit les index (ex: 15) en lettres Excel (ex: "P")
    lettre_n1, lettre_n2 = xl_col_to_name(idx_n1), xl_col_to_name(idx_n2)
    lettre_n3, lettre_n4 = xl_col_to_name(idx_n3), xl_col_to_name(idx_n4)
    
    lettre_g1 = xl_col_to_name(idx_g1)
    lettre_g2, lettre_g3, lettre_g4 = xl_col_to_name(idx_g2), xl_col_to_name(idx_g3), xl_col_to_name(idx_g4)

    # 5) Menu déroulants (Data Validation en cascade et simples)
    # N1 : Menu fixe
    worksheet.data_validation(1, idx_n1, max_row, idx_n1, {'validate': 'list', 'source': '=N1_Liste'})
    
    # N2, N3, N4 : Menus dépendants de la colonne précédente (INDIRECT + SUBSTITUTE pour les espaces)
    worksheet.data_validation(1, idx_n2, max_row, idx_n2, {'validate': 'list', 'source': f'=INDIRECT(SUBSTITUTE(${lettre_n1}2, " ", "_"))'})
    worksheet.data_validation(1, idx_n3, max_row, idx_n3, {'validate': 'list', 'source': f'=INDIRECT(SUBSTITUTE(${lettre_n2}2, " ", "_"))'})
    worksheet.data_validation(1, idx_n4, max_row, idx_n4, {'validate': 'list', 'source': f'=INDIRECT(SUBSTITUTE(${lettre_n3}2, " ", "_"))'})

    menus_simples_actifs = {
        "Discipline": ["Economie", "Gestion", "Droit & Science Politique", "A vérifier"],
        "Typologie_Nationalite": ["Française", "UE", "Hors UE", "Non renseigné"],
        "Typologie_Etablissement": ["Université Etrangère - Pays UE", "Université Etrangère - Hors UE", "Université Côte d'Azur", "Ecole de Commerce / Ingénieur", "Autre Université Française", "Non renseigné (Erreur ou NaN)"],
        "Césure": ["Oui", "Non", "A vérifier"],
        "Civilité" : ["Monsieur", "Madame"]
    }
    # On extrait la liste des labos directement depuis les données !
    if 'Unité de recherche (libellé abrégé)' in df.columns:
        menus_simples_actifs["Unité de recherche (libellé abrégé)"] = sorted([str(x) for x in df['Unité de recherche (libellé abrégé)'].unique() if str(x) not in ["A vérifier", "nan", "None", ""]]) + ["A vérifier"]

    for nom_colonne, liste_choix in menus_simples_actifs.items():
        if nom_colonne in df.columns:
            idx = df.columns.get_loc(nom_colonne)
            worksheet.data_validation(1, idx, max_row, idx, {
                'validate': 'list',
                'source': liste_choix
            })

    # 6) Mise en forme conditionnelle (Alerte Globale + A Vérifier)
    # Les designs (Rouge PÉTANT pour l'erreur, Jaune pour A vérifier)
    format_rouge = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF', 'bold': True}) # Fond rouge vif, texte blanc gras
    format_jaune = workbook.add_format({'bg_color': "#FFEE00", 'font_color': '#000000', 'bold': True})
    # Règle 1: le rouge (Incohérence Globale sur toute la ligne)
    err_n2 = f'ISERROR(MATCH(${lettre_n2}2, INDIRECT(SUBSTITUTE(${lettre_n1}2, " ", "_")), 0))'
    err_n3 = f'ISERROR(MATCH(${lettre_n3}2, INDIRECT(SUBSTITUTE(${lettre_n2}2, " ", "_")), 0))'
    err_n4 = f'ISERROR(MATCH(${lettre_n4}2, INDIRECT(SUBSTITUTE(${lettre_n3}2, " ", "_")), 0))'
    
    chaine_cassee = f'OR({err_n2}, {err_n3}, {err_n4})'
    ligne_touchee = f'OR(${lettre_n1}2<>${lettre_g1}2, ${lettre_n2}2<>${lettre_g2}2, ${lettre_n3}2<>${lettre_g3}2, ${lettre_n4}2<>${lettre_g4}2)'

    formule_globale = f'=AND({chaine_cassee}, {ligne_touchee})'

    # On applique de la colonne 0 à la dernière colonne (max_col - 1)
    worksheet.conditional_format(1, 0, max_row, max_col - 1, {
        'type': 'formula', 
        'criteria': formule_globale, 
        'format': format_rouge
    })

    # Règle 2 : du jaune sur tt la ligne si "A vérifier" est présent quelque part (dans n'importe quelle colonne visible)
    # Cette formule Excel compte si le mot "A vérifier" apparaît n'importe où sur la ligne
    derniere_lettre_visible = xl_col_to_name(nb_colonnes_visibles - 1)
    formule_jaune = f'=COUNTIF($A2:${derniere_lettre_visible}2, "A vérifier")>0'
    worksheet.conditional_format(1, 0, max_row, max_col - 1, {'type': 'formula', 'criteria': formule_jaune, 'format': format_jaune})

    writer.close()


# ===================================================================
# 5) MOTEUR DE CALCUL DES KPIS (CUBE DE DONNÉES POWER BI)
# ===================================================================

def calculer_kpis_cohorte(df_1, df_2, df_3, df_4, annee_resultat, nom_disc, nom_labo):
    """
    Calcule les KPIs pour un sous-groupe spécifique (filtré en amont)
    et génère la clé primaire pour Power BI.
    """
    len_1 = len(df_1)
    len_4 = len(df_4)
    
    # KPIs Nouveaux entrants (Base 1)
    taux_fi = (df_1['Financement_N1'] == "Financé").mean() if len_1 > 0 else 0
    taux_inter = df_1['Typologie_Nationalite'].isin(["UE", "Hors UE"]).mean() if len_1 > 0 else 0
    taux_attract = (df_1['Typologie_Etablissement'] != "Université Côte d'Azur").mean() if len_1 > 0 else 0
    nb_cifre = (df_1['Financement_N4'] == "CIFRE").sum() if len_1 > 0 else 0
    taux_fem_nouveaux = (df_1['Civilité'].str.lower().str.contains('mme|madame|f', na=False)).mean() if len_1 > 0 else 0
    
    # Féminisation Globale (Base 4)
    taux_fem = (df_4['Civilité'].str.lower().str.contains('mme|madame|f', na=False)).mean() if len_4 > 0 else 0
    
    # Fuite Précoce (Base 2)
    base_abandons = df_2[df_2['Type_Depart'] != 'Délai maximum atteint'] if not df_2.empty else df_2
    taux_fuite = (base_abandons['Anciennete_Depart'] == 'Avant 3 ans (inclus)').mean() if not base_abandons.empty else 0
    nb_delais_depasses = (df_2['Type_Depart'] == 'Délai maximum atteint').sum() if not df_2.empty else 0

    # KPIs Sorties (Base 3)
    duree_col = pd.to_numeric(df_3["Durée sans suspension ni césure (mois)"], errors='coerce').dropna() if not df_3.empty else pd.Series(dtype=float)
    duree_mediane = duree_col.median() if not duree_col.empty else 0
    taux_depassement = (duree_col > 72).mean() if not duree_col.empty else 0
    taux_cotutelle = df_3['Cotutelle_Internationale'].isin(["UE", "Hors UE"]).mean() if not df_3.empty else 0
    taux_fem_soutenances = (df_3['Civilité'].str.lower().str.contains('mme|madame|f', na=False)).mean() if not df_3.empty else 0

    # CRÉATION DE LA CLÉ PRIMAIRE (La magie pour Power BI)
    cle_filtre = f"{nom_disc}-{nom_labo}"

    return {
        "Année_Universitaire": int(annee_resultat),
        "Discipline": nom_disc,
        "Laboratoire": nom_labo,
        "Clé_filtre": cle_filtre,
        "Total_Inscrits_premiere_annee": len_1,
        "Total_Inscrits": len_4,
        "Total_Abandons": len(base_abandons),
        "Nb_Delais_Depasses": int(nb_delais_depasses),
        "Total_Soutenances": len(df_3),
        "Taux_Financement_Nouveaux": round(taux_fi, 4),
        "Taux_Internationalisation": round(taux_inter, 4),
        "Taux_Attractivite_Exterieure": round(taux_attract, 4),
        "Taux_Feminisation": round(taux_fem, 4),
        "Taux_Feminisation_Nouveaux": round(taux_fem_nouveaux, 4),
        "Taux_Feminisation_Soutenances": round(taux_fem_soutenances, 4),
        "Nb_CIFRE": int(nb_cifre),
        "Duree_Mediane_Soutenances": round(duree_mediane, 1),
        "Taux_Depassement_Duree": round(taux_depassement, 4),
        "Taux_Cotutelle": round(taux_cotutelle, 4),
        "Taux_Fuite_Precoce": round(taux_fuite, 4)
    }