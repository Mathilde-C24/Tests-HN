# Extracteur de Correspondance Historique

Outil d'extraction automatique de métadonnées depuis des correspondances historiques au format DOCX vers CSV.

## 📋 Description

Ce script Python extrait automatiquement les informations suivantes depuis des fichiers Word (.docx) contenant des correspondances historiques :

- **Numéro de lettre** (numérotation automatique)
- **Auteur** (expéditeur)
- **Destinataire**
- **Date** (normalisée au format AAAA-MM-JJ ou AN-mois_révolutionnaire-JJ)
- **Lieu** (si présent)
- **Nombre de mots** (sans l'en-tête)

## 🎯 Cas d'usage

Conçu pour le projet de correspondance de Saint-Domingue (1760-1804), ce script peut être adapté à toute collection de lettres historiques structurées de manière similaire.

## 📦 Installation

### Prérequis

- Python 3.7+
- pip

### Dépendances

```bash
pip install -r requirements.txt
```

## 🚀 Utilisation

### Utilisation basique

Placez vos fichiers .docx dans un dossier et lancez :

```bash
python extracteur_correspondance.py /chemin/vers/dossier
```

Si aucun chemin n'est spécifié, le script utilise le dossier courant.

### Structure attendue des fichiers

Les fichiers .docx doivent contenir :
- **En-têtes** en style "Heading 1" au format : `Auteur à Destinataire`
- **Date et lieu** sur la ligne suivant l'en-tête (ou après une ligne vide)
- **Contenu** de la lettre dans les paragraphes suivants

### Formats de dates supportés

#### Calendrier grégorien
- `12 juin 1778`
- `12 janv. 1778`
- `12 7bre 1778` (septembre)
- `12 xbre 1778` (décembre)

#### Calendrier révolutionnaire
- `3 frimaire an VIII`
- `15 thermidor an 2`

**Note :** Si une date grégorienne et révolutionnaire sont présentes, la grégorienne est prioritaire.

## 📄 Format de sortie

Le script génère un fichier CSV avec les colonnes suivantes :

```csv
numero;auteur;destinataire;date;lieu;nb_mots
1;Marie Dupont;Jean Martin;1778-06-12;Les Cayes;245
2;Jean Martin;Marie Dupont;VIII-frimaire-03;;312
```

- **Délimiteur** : `;` (point-virgule)
- **Encodage** : UTF-8 avec BOM
- **Format de date** : 
  - Grégorien : `AAAA-MM-JJ`
  - Révolutionnaire : `AN-mois_révolutionnaire-JJ`

## 🔧 Configuration

Les noms de fichiers à traiter sont définis dans la constante `FICHIERS` (lignes 16-26). Modifiez cette liste selon vos besoins.

## 📚 Structure du projet

```
.
├── extracteur_correspondance.py  # Script principal
├── requirements.txt               # Dépendances Python
├── README.md                      # Cette documentation
├── LICENSE                        # Licence MIT
└── exemple/                       # Exemples (optionnel)
    ├── input/                     # Fichiers .docx exemples
    └── output/                    # CSV résultant
```

## 🤝 Contribution

Les contributions sont bienvenues ! N'hésitez pas à :
- Signaler des bugs
- Proposer des améliorations
- Soumettre des pull requests

## 📝 Licence

Ce projet est sous licence MIT. Voir le fichier [LICENSE](LICENSE) pour plus de détails.

## 👤 Auteur

**Mathilde Chollet**
- Master 1 Humanités Numériques
- Généré avec l'assistance de Claude (Anthropic)

## 🙏 Remerciements

- Projet de recherche sur la correspondance de Saint-Domingue
- Classiques Garnier (édition des correspondances)

## 📊 Métadonnées FAIR

- **Version** : 1.0.0
- **Date de création** : 2024-12-07
- **Langage** : Python 3.7+
- **Format d'entrée** : DOCX (Microsoft Word)
- **Format de sortie** : CSV (UTF-8)
- **Standards** : Principes FAIR (Findable, Accessible, Interoperable, Reusable)

## 🔗 Ressources

- [Documentation python-docx](https://python-docx.readthedocs.io/)
- [Calendrier révolutionnaire français](https://fr.wikipedia.org/wiki/Calendrier_r%C3%A9publicain)
