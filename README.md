# Suivi de la performance de prestataires de relation client — Excel & Power BI

## Objectif

Consolider et piloter la performance de **4 prestataires externes** de relation client à partir de leurs exports bruts respectifs, hétérogènes dans leur format — exercice qui reproduit directement le pilotage multi-prestataires d'un service comme le DPEX chez EDF (suivi budgétaire, indicateurs métier, facturation, qualité).

Le projet couvre trois volets : **Excel** (harmonisation, formules complexes, TCD, macros VBA), **modélisation en étoile**, et **Power BI** (dashboard interactif).

## Sources de données

4 fichiers CSV, un par prestataire, générés pour simuler des systèmes de reporting différents (mêmes concepts métier, formats et nommages différents) — environ 9 600 lignes cumulées.

### Prestataire A
Colonnes : `ID_Appel`, `Date_Appel` (JJ/MM/AAAA), `Agent`, `Duree_Traitement_Sec`, `Statut`, `Note_Satisfaction` (/5), `Montant_Facture_EUR`, `Canal`, `Sujet`, `Appel_Traite`, `Appel_Repris`, `Conformite`, `NPS` (/10), `Reclamation`

### Prestataire B
Colonnes : `IdTicket`, `DateTicket` (AAAA-MM-JJ), `Operateur`, `TempsTraitement_Min`, `Resolu`, `Satisfaction_Client`, `CoutFacturation`, `Canal_Contact`, `Motif`, `Traite`, `Reprise_Appel`, `Conforme`, `Score_NPS`, `Motif_Reclamation`

### Prestataire C
Colonnes : `Reference_Dossier`, `Jour` (numéro de série Excel), `Nom_Agent`, `Duree_Min`, `Dossier_Cloture`, `Note_Client` (/5), `Prix_Prestation`, `Type_Contact`, `Objet`, `Traite`, `Repris`, `Conformite_Controle`, `NPS_Client` (/10), `Est_Reclamation`

### Prestataire D
Colonnes (en anglais) : `N_Contact`, `Date` (MM/JJ/AAAA), `Agent_ID`, `Handling_Time_Sec`, `Status`, `Rating` (/5), `Billed_Amount`, `Channel`, `Topic`, `Handled`, `Repeat_Call`, `Compliance`, `NPS_Score` (/10), `Complaint`

## Correspondance des concepts entre prestataires (à harmoniser)

| Concept métier | Prestataire A | Prestataire B | Prestataire C | Prestataire D |
|---|---|---|---|---|
| Identifiant | ID_Appel | IdTicket | Reference_Dossier | N_Contact |
| Date | Date_Appel | DateTicket | Jour | Date |
| Agent | Agent | Operateur | Nom_Agent | Agent_ID |
| Durée de traitement | Duree_Traitement_Sec | TempsTraitement_Min | Duree_Min | Handling_Time_Sec |
| Résolution | Statut | Resolu | Dossier_Cloture | Status |
| Satisfaction | Note_Satisfaction | Satisfaction_Client | Note_Client | Rating |
| Montant facturé | Montant_Facture_EUR | CoutFacturation | Prix_Prestation | Billed_Amount |
| Canal | Canal | Canal_Contact | Type_Contact | Channel |
| Sujet | Sujet | Motif | Objet | Topic |
| Appel traité | Appel_Traite | Traite | Traite | Handled |
| Appel repris | Appel_Repris | Reprise_Appel | Repris | Repeat_Call |
| Conformité | Conformite | Conforme | Conformite_Controle | Compliance |
| NPS | NPS | Score_NPS | NPS_Client | NPS_Score |
| Réclamation | Reclamation | Motif_Reclamation | Est_Reclamation | Complaint |

## Synthèse nettoyage 

## Table Prestataire A
1. Renommage des colonnes et harmonisation des libellés en français.
2. Nettoyage des valeurs aberrantes dans `Duree_Traitement_Sec` :
   - Remplacement des valeurs < 10s et > 2000s par `null`.
3. Correction des types de données :
   - `Note_Satisfaction_Client` : String → Int
   - `NPS` : String → Int
   - `Montant_Facture` : String → Float
4. Ajout d'une colonne `ID_Agent` basée sur le nom de l’agent.
5. Remplacement des champs vides dans les colonnes non quantitatives par `Non renseigné`.

## Table Prestataire B
1. Renommage des colonnes et harmonisation des libellés en français.
2. Création de la colonne `Duree_Traitement_Sec` :
   - Conversion des minutes en secondes.
3. Correction des types de données :
   - `Note_Satisfaction_Client` : String → Int
   - `NPS` : String → Int
   - `Montant_Facture` : String → Float
4. Normalisation des doublons dans `Canal` :
   - `Tel` → `Téléphone`
   - `Mail` → `Email`
5. Conversion des colonnes binaires en valeurs textuelles :
   - `Oui` / `Non`
6. Remplacement des champs vides dans les colonnes non quantitatives par `Non renseigné`.
7. Ajout d'une colonne Réclammation

## Table Prestataire C

1. Renommage des colonnes et harmonisation des libellés en français afin d’aligner la structure avec les tables A et B.

2. Normalisation des identifiants agents :
   - Conversion des formats `AGT-01`, `AGT-02`, etc. vers un format standardisé `AG001`, `AG002`, etc.

3. Correction des types de données :
   - Conversion de la colonne `Date` depuis un format texte (String) vers un type Date.
   - Renommage de la colonne en `Date` après conversion.

4. Création de la colonne `Duree_Traitement_Sec` :
   - Transformation de la durée exprimée en minutes vers une durée en secondes.
   - Formule : `Durée (min) × 60`.

5. Remplacement des champs vides dans les colonnes non quantitatives par la valeur `Non renseigné` afin d’assurer une cohérence des données.

6. Conversion des colonnes binaires en valeurs textuelles :
   - Transformation des valeurs numériques ou booléennes en `Oui` / `Non`.

7. Normalisation des valeurs booléennes textuelles :
   - Remplacement des valeurs `VRAI`, `FAUSSE`, `TRUE`, `FALSE` par `Oui` / `Non` dans les colonnes `Traite`, `Reclamation` et `Dossier_Cloture`.

8. Nettoyage des valeurs aberrantes dans `Duree_Traitement_Sec` :
   - Remplacement des valeurs < 10 secondes par `null`.

9. Nettoyage des valeurs aberrantes dans `Prix_Prestation` :
   - Remplacement des valeurs < 12 € par `null`.

## Table Prestataire D

1. Renommage et traduction des colonnes :
   - Traduction des libellés anglais en français pour harmoniser la structure avec les autres prestataires.
   - Exemple : `Social Media` → `Réseaux_Sociaux`.

2. Normalisation des identifiants agents :
   - Conversion des formats `AG101`, `AG102`, etc. vers un format standardisé `AG001`, `AG002`, etc.

3. Correction des types de données :
   - Conversion des colonnes `Note_Satisfaction_Client` et `NPS` depuis le type String vers le type Float.

4. Correction du format de date :
   - Détection et inversion des dates mal formatées (jour/mois inversés).
   - Création d’une nouvelle colonne `Date_Corrigee` :
     - Si le mois > jour et jour < mois, inversion automatique des positions jour/mois.
     - Format final : `JJ/MM/AAAA`.

5. Nettoyage des valeurs aberrantes dans `Temps_Traitement_Sec` :
   - Remplacement des valeurs < 10 secondes par `null`.

6. Conversion des colonnes binaires en valeurs textuelles :
   - Transformation des valeurs numériques ou booléennes en `Oui` / `Non`.

## Modèle en étoile 
- **Dimensions  : `Dim_Agent`, `Dim_Agent`, `Dim_Prestataire`
- **Tables faits : `Appels_Global`

## Plan de travail

### Excel
1. Exploration de chaque fichier séparément, confirmation des problèmes listés ci-dessus
2. Nettoyage individuel de chaque source (Power Query)
3. Harmonisation : renommage des colonnes vers un nom commun, conversion des unités (durée), conversion des formats de date, uniformisation des booléens et de l'échelle NPS
4. Fusion des 4 sources en une table unique (`Append`), avec colonne `Prestataire`
5. Construction du modèle en étoile
6. Formules complexes pour les KPI 
7. Tableaux croisés dynamiques sur plusieurs angles (prestataire, mois, canal, sujet)
8. Macros VBA :
   - rafraîchissement des données
   - contrôle qualité (ex : alerte si le taux de non-conformité ou de valeurs manquantes dépasse un seuil sur un prestataire)
   - génération automatique d'un rapport de synthèse mensuel

### Power BI
9. Import du modèle construit (ou reconstruction directe dans Power BI)
10. Dashboard interactif comparant la performance des 4 prestataires

### Documentation
11. Mise à jour de ce README au fil du projet

##  Familles de KPI contrôlées

Quatre familles de KPI ont été analysées afin d’évaluer la performance globale des prestataires :

### 1️ KPI de qualité
Mesurent la conformité et la fiabilité du traitement des appels.  
**Indicateurs inclus :**
- Nombre d’appels reçus  
- Nombre d’appels traités  
- Nombre d’appels contrôlés  
- Taux de conformité  
- Taux d’erreur  
- Taux de reprises  
- Taux de conformité au premier passage  
- Taux de réclamation  

### 2 KPI de délai
Évaluent la rapidité de prise en charge et de résolution.  


### 3️ KPI de satisfaction client
Mesurent la perception du service par les clients.  


### 4️ KPI de productivité
Apprécient la capacité des équipes à absorber la charge de travail.  



## Stack

- Excel (Power Query, formules, TCD, VBA)
- Power BI (modélisation, DAX, dashboard)

## Structure du repo 

```
├── data/
│   └── brut/
│   │    ├── Prestataire_A.csv
│   │    ├── Prestataire_B.csv
│   │    ├── Prestataire_C.csv
│   │    └── Prestataire_D.csv
│   └── corrigé/
│        ├── Prestataire_A.corrigé.csv
│        ├── Prestataire_B.corrigé.csv
│        ├── Prestataire_C.corrigé.csv
│        └── Prestataire_D.corrigé.csv
├── excel/
│   └── suivi_prestataires.xlsx
├── powerbi/
│   └── suivi_prestataires.pbix
├── docs/
│   └── captures/
└── README.md
```
