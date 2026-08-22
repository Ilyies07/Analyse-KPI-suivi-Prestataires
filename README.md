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

# Table Prestataire A
1. Renommage des colonnes et des champs des colonnes en français pour mieux harmoniser les données 
2. Traitement des valeurs aberrantes dans Duree traitement seconde en les remplaçants pas null ( les valeeurs > 2000 et < 0)
3. Correction des types Note satisfaction client, NPS (String -> INT) Montant_Facture String -> Float
4. Ajout d'une colonne Id Agent 
5. Remplacement de certains champs vide par Non Renseigné dans les colonnes non quantitatives.


## Modèle en étoile (à construire)

## Plan de travail

### Excel
1. Exploration de chaque fichier séparément, confirmation des problèmes listés ci-dessus
2. Nettoyage individuel de chaque source (Power Query)
3. Harmonisation : renommage des colonnes vers un nom commun, conversion des unités (durée), conversion des formats de date, uniformisation des booléens et de l'échelle NPS
4. Fusion des 4 sources en une table unique (`Append`), avec colonne `Prestataire`
5. Construction du modèle en étoile
6. Formules complexes pour les KPI (à définir toi-même — voir familles ci-dessous)
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

## Familles de KPI à définir et calculer toi-même

- **Volumétrie** : nombre de contacts reçus, traités, non traités par prestataire
- **Qualité** : taux de résolution, taux de conformité, taux de réclamation, taux de reprise (rappels)
- **Délai** : temps de traitement moyen/médian par prestataire, par canal
- **Satisfaction** : note moyenne, NPS moyen, répartition des notes
- **Budgétaire** : montant facturé total et moyen par prestataire, par sujet, par mois
- **Comparatif** : classement des prestataires sur un ou plusieurs indicateurs combinés

## Stack

- Excel (Power Query, formules, TCD, VBA)
- Power BI (modélisation, DAX, dashboard)

## Structure du repo (à adapter)

```
├── data/
│   └── raw/
│       ├── Prestataire_A.csv
│       ├── Prestataire_B.csv
│       ├── Prestataire_C.csv
│       └── Prestataire_D.csv
├── excel/
│   └── suivi_prestataires.xlsx
├── powerbi/
│   └── suivi_prestataires.pbix
├── docs/
│   └── captures/
└── README.md
```

## Notes

*(à compléter au fil du projet : difficultés de réconciliation entre prestataires, choix de modélisation, macros VBA développées, indicateurs et insights finalement retenus)*
