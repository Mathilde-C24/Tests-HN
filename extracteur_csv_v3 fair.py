"""
Extracteur de correspondance vers CSV
Extrait : N° lettre, AUTEUR, DESTINATAIRE, DATE, LIEU, NB_MOTS

Auteur : Mathilde Chollet
Date : 2024-12-07
Version : 1.0.0
Généré avec l'assistance de Claude (Anthropic)
Licence : MIT
"""

import csv
import re
import sys
from docx import Document
from pathlib import Path

# Configuration des fichiers
FICHIERS = [
    "Lettres 1760-1764.docx",
    "Lettres 1765-1769.docx",
    "Lettres 1770-1774.docx",
    "Lettres 1775-1779.docx",
    "Lettres 1780-1784.docx",
    "Lettres 1785-1789.docx",
    "Lettres 1790-1794.docx",
    "Lettres 1795-1799.docx",
    "Lettres 1800-1804.docx"
]

# Chemin par défaut (modifiable via argument)
DOSSIER_DEFAUT = "."

# Mois grégoriens (orthographe ancienne acceptée)
MOIS_GREGORIENS = {
    'janvier': 1, 'janv.': 1, 'janv': 1,
    'février': 2, 'fevrier': 2, 'fév.': 2, 'fev.': 2, 'fev': 2,
    'mars': 3,
    'avril': 4, 'avr.': 4, 'avr': 4,
    'mai': 5,
    'juin': 6,
    'juillet': 7, 'juil.': 7, 'juil': 7,
    'août': 8, 'aoust': 8, 'aout': 8,
    'septembre': 9, '7bre': 9, 'sept.': 9, 'sept': 9, '7bre.': 9,
    'octobre': 10, '8bre': 10, 'oct.': 10, 'oct': 10, '8bre.': 10,
    'novembre': 11, '9bre': 11, 'nov.': 11, 'nov': 11, '9bre.': 11,
    'décembre': 12, 'decembre': 12, '10bre': 12, 'xbre': 12, 'déc.': 12, 'dec.': 12, '10bre.': 12, 'xbre.': 12
}

# Mois révolutionnaires
MOIS_REVOLUTIONNAIRES = [
    'vendémiaire', 'vendemiaire',
    'brumaire',
    'frimaire',
    'nivôse', 'nivose',
    'pluviôse', 'pluviose',
    'ventôse', 'ventose',
    'germinal',
    'floréal', 'floreal',
    'prairial',
    'messidor',
    'thermidor',
    'fructidor'
]

def extraire_correspondants(entete):
    """Extrait l'expéditeur et le destinataire."""
    entete = re.sub(r'^\d+\.\s*', '', entete)
    
    if ' à ' in entete:
        parties = entete.split(' à ', 1)
        expediteur = parties[0].strip()
        destinataire = parties[1].split(' - ')[0].strip() if ' - ' in parties[1] else parties[1].strip()
        return expediteur, destinataire
    
    return None, None

def detecter_date_revolutionnaire(ligne):
    """Détecte si la ligne contient une date révolutionnaire."""
    ligne_lower = ligne.lower()
    
    # Chercher un mois révolutionnaire
    for mois_revo in MOIS_REVOLUTIONNAIRES:
        if mois_revo in ligne_lower:
            # Chercher le jour et l'année
            # Format : "3 frimaire an VIII" ou "le 3 frimaire an 8"
            pattern = r'(\d{1,2})\s+' + mois_revo + r'\s+(?:an\s+)?([IVXLCDM]+|\d{1,2})'
            match = re.search(pattern, ligne_lower, re.IGNORECASE)
            if match:
                jour = int(match.group(1))
                an = match.group(2)
                # Nettoyer le nom du mois (enlever accents pour uniformiser)
                mois_clean = mois_revo.replace('é', 'e').replace('ô', 'o')
                if mois_clean.endswith('e'):
                    mois_clean = mois_clean[:-1] + 'e'
                return f"{an}-{mois_revo}-{jour:02d}"
    
    return None

def extraire_date_lieu(ligne):
    """
    Extrait la date normalisée (AAAA-MM-JJ ou AN-MOIS_REVO-JJ) et le lieu.
    Priorité : date grégorienne si présente avec date révolutionnaire.
    """
    ligne = ligne.strip()
    
    # Retirer les crochets
    ligne_clean = re.sub(r'\[([^\]]+)\]', r'\1', ligne)
    
    # Chercher une date grégorienne
    # Pattern : jour mois année, avec possibilité de lieu avant
    # Ex: "Les Cayes 12 juin 1778" ou "12 juin 1778"
    pattern_greg = r'(\d{1,2}(?:er)?)\s+([a-zàéèêôû0-9]+\.?)\s+(\d{4})'
    match_greg = re.search(pattern_greg, ligne_clean, re.IGNORECASE)
    
    date_normalisee = None
    lieu = None
    
    if match_greg:
        jour_str, mois_str, annee = match_greg.groups()
        jour = int(jour_str.replace('er', ''))
        mois_lower = mois_str.lower()
        
        if mois_lower in MOIS_GREGORIENS:
            mois = MOIS_GREGORIENS[mois_lower]
            date_normalisee = f"{annee}-{mois:02d}-{jour:02d}"
            
            # Extraire le lieu (texte avant la date)
            debut_date = match_greg.start()
            if debut_date > 0:
                lieu_brut = ligne_clean[:debut_date].strip()
                # Nettoyer les mots comme "À", "à", "le", etc.
                lieu = re.sub(r'^(à|À|le|Le|ce|Ce)\s+', '', lieu_brut).strip()
                if lieu and len(lieu) > 1:
                    lieu = lieu
                else:
                    lieu = None
    
    # Si pas de date grégorienne, chercher date révolutionnaire
    if not date_normalisee:
        date_normalisee = detecter_date_revolutionnaire(ligne_clean)
        
        # Si date révolutionnaire trouvée, essayer d'extraire le lieu
        if date_normalisee:
            # Le lieu est probablement au début de la ligne
            mots = ligne_clean.split()
            if len(mots) > 0 and not any(m in mots[0].lower() for m in ['le', 'ce', 'du']):
                # Premier mot pourrait être un lieu
                if not any(c.isdigit() for c in mots[0]):
                    lieu = mots[0].strip(',')
    
    return date_normalisee, lieu

def compter_mots(texte):
    """Compte le nombre de mots dans un texte."""
    if not texte:
        return 0
    # Découper sur les espaces et compter
    mots = texte.split()
    return len(mots)

def extraire_lettres_du_document(chemin_docx, numero_debut=1):
    """Extrait toutes les lettres d'un fichier .docx."""
    doc = Document(chemin_docx)
    lettres = []
    lettre_courante = None
    contenu_paragraphes = []
    date_trouvee = False
    ligne_vide_apres_entete = False
    numero_courant = numero_debut
    
    for para in doc.paragraphs:
        texte = para.text.strip()
        
        # Détecter un en-tête de lettre (style Heading 1)
        if para.style.name == 'Heading 1':
            # Sauvegarder la lettre précédente
            if lettre_courante and contenu_paragraphes:
                contenu = ' '.join(contenu_paragraphes)
                lettre_courante['nb_mots'] = compter_mots(contenu)
                lettres.append(lettre_courante)
            
            # Commencer une nouvelle lettre
            expediteur, destinataire = extraire_correspondants(texte)
            numero = numero_courant
            numero_courant += 1
            
            lettre_courante = {
                'numero': numero,
                'auteur': expediteur,
                'destinataire': destinataire,
                'date': None,
                'lieu': None,
                'nb_mots': 0
            }
            contenu_paragraphes = []
            date_trouvee = False
            ligne_vide_apres_entete = False
            continue
        
        # Si on est dans une lettre
        if lettre_courante is not None:
            # Gérer les lignes vides après l'en-tête
            if not texte and not date_trouvee and not contenu_paragraphes:
                ligne_vide_apres_entete = True
                continue
            
            # Si on n'a pas encore trouvé la date, vérifier si c'est la ligne de date
            if not date_trouvee and not contenu_paragraphes:
                date, lieu = extraire_date_lieu(texte)
                if date:
                    lettre_courante['date'] = date
                    lettre_courante['lieu'] = lieu
                    date_trouvee = True
                    continue
            
            # Sinon, c'est du contenu
            if texte:
                contenu_paragraphes.append(texte)
    
    # Sauvegarder la dernière lettre
    if lettre_courante and contenu_paragraphes:
        contenu = ' '.join(contenu_paragraphes)
        lettre_courante['nb_mots'] = compter_mots(contenu)
        lettres.append(lettre_courante)
    
    return lettres

def traiter_tous_les_fichiers(dossier_source):
    """Traite tous les fichiers .docx et crée un CSV."""
    dossier = Path(dossier_source)
    toutes_lettres = []
    
    for nom_fichier in FICHIERS:
        chemin_fichier = dossier / nom_fichier
        
        if not chemin_fichier.exists():
            print(f"⚠️  Fichier non trouvé : {nom_fichier} (ignoré)")
            continue
        
        print(f"📖 Traitement de {nom_fichier}...")
        lettres = extraire_lettres_du_document(chemin_fichier, numero_debut=len(toutes_lettres) + 1)
        toutes_lettres.extend(lettres)
        print(f"   ✓ {len(lettres)} lettres extraites")
    
    print(f"\n📊 Résumé :")
    print(f"   • Total de lettres : {len(toutes_lettres)}")
    lettres_datees = [l for l in toutes_lettres if l['date']]
    print(f"   • Lettres datées : {len(lettres_datees)}")
    
    # Sauvegarder en CSV
    chemin_csv = dossier / 'correspondance_st-domingue.csv'
    with open(chemin_csv, 'w', newline='', encoding='utf-8-sig') as f:
        fieldnames = ['numero', 'auteur', 'destinataire', 'date', 'lieu', 'nb_mots']
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
        
        writer.writeheader()
        writer.writerows(toutes_lettres)
    
    print(f"\n✅ Fichier CSV créé : {chemin_csv}")
    return toutes_lettres

# Utilisation
if __name__ == "__main__":
    # Utiliser argument en ligne de commande ou dossier courant
    if len(sys.argv) > 1:
        dossier_source = sys.argv[1]
    else:
        dossier_source = DOSSIER_DEFAUT
    
    print("🚀 Extraction de la correspondance de Saint-Domingue\n")
    print(f"📁 Dossier source : {dossier_source}\n")
    
    resultat = traiter_tous_les_fichiers(dossier_source)
    
    print("\n🎉 Extraction terminée !")
