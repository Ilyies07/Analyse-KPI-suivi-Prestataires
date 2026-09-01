#  Suivi & Performance des Prestataires de Relation Client (Excel & Power BI)

##  1. Objectif du Projet

Ce projet vise à **consolider et piloter la performance de 4 prestataires externes** de relation client (Alpha Services, Beta Solutions, Cobalt Support, Delta Connect) à partir d'exports bruts hétérogènes. 

Cette démarche reproduit fidèlement le pilotage multi-prestataires d'un service d'exploitation (type DPEX chez EDF) en couvrant :
* Le suivi budgétaire et de facturation
* Le contrôle qualité et la conformité
* La productivité et les délais de traitement
* La satisfaction client (CSAT & NPS)

Le projet s'articule autour de trois piliers principaux :
1. **Excel & Power Query :** Nettoyage, harmonisation, calcul de KPI complexes, TCD et automatisation par Macros VBA.
2. **Modélisation de Données :** Conception d'un schéma en étoile (*Star Schema*).
3. **Power BI :** Dashboard analytique et interactif de pilotage.

---

##  2. Sources de Données & Harmonisation

Le jeu de données comprend **4 fichiers CSV** (environ 9 600 lignes cumulées) simulant des systèmes de reporting d'entreprises distinctes avec des formats, unités et nommages variables.

###  Table de Correspondance des Concepts

| Concept Métier | Prestataire A | Prestataire B | Prestataire C | Prestataire D |
| :--- | :--- | :--- | :--- | :--- |
| **Identifiant** | `ID_Appel` | `IdTicket` | `Reference_Dossier` | `N_Contact` |
| **Date** | `Date_Appel` | `DateTicket` | `Jour` | `Date` |
| **Agent** | `Agent` | `Operateur` | `Nom_Agent` | `Agent_ID` |
| **Durée de traitement**| `Duree_Traitement_Sec` | `TempsTraitement_Min` | `Duree_Min` | `Handling_Time_Sec` |
| **Résolution** | `Statut` | `Resolu` | `Dossier_Cloture` | `Status` |
| **Satisfaction** | `Note_Satisfaction` | `Satisfaction_Client` | `Note_Client` | `Rating` |
| **Montant facturé** | `Montant_Facture_EUR` | `CoutFacturation` | `Prix_Prestation` | `Billed_Amount` |
| **Canal** | `Canal` | `Canal_Contact` | `Type_Contact` | `Channel` |
| **Sujet** | `Sujet` | `Motif` | `Objet` | `Topic` |
| **Appel traité** | `Appel_Traite` | `Traite` | `Traite` | `Handled` |
| **Appel repris** | `Appel_Repris` | `Reprise_Appel` | `Repris` | `Repeat_Call` |
| **Conformité** | `Conformite` | `Conforme` | `Conformite_Controle` | `Compliance` |
| **NPS** | `NPS` | `Score_NPS` | `NPS_Client` | `NPS_Score` |
| **Réclamation** | `Reclamation` | `Motif_Reclamation` | `Est_Reclamation` | `Complaint` |

---

##  3. Synthèse du Nettoyage des Données (Power Query)

###  Table Prestataire A
1. Renommage des colonnes et harmonisation des libellés en français.
2. Nettoyage des valeurs aberrantes dans `Duree_Traitement_Sec` (remplacement des valeurs < 10s et > 2000s par `null`).
3. Typpage des données : `Note_Satisfaction_Client` (Int), `NPS` (Int), `Montant_Facture` (Float).
4. Création de la colonne `ID_Agent` basée sur le nom de l’agent.
5. Remplacement des valeurs vides non quantitatives par `Non renseigné`.

###  Table Prestataire B
1. Renommage et harmonisation des colonnes.
2. Conversion de la durée de minutes en secondes (`Duree_Traitement_Sec`).
3. Normalisation des doublons dans `Canal` (`Tel` → `Téléphone`, `Mail` → `Email`).
4. Conversion des indicateurs binaires en valeurs textuelles (`Oui` / `Non`).
5. Ajout et structuration de la colonne `Réclamation`.

###  Table Prestataire C
1. Normalisation des identifiants agents (`AGT-01` → `AG001`).
2. Conversion de la colonne `Date` (type Date à partir du numéro de série Excel).
3. Conversion de la durée de minutes en secondes (`Durée min × 60`).
4. Normalisation des booléens textuels (`VRAI`, `FAUSSE`, `TRUE`, `FALSE` → `Oui` / `Non`).
5. Traitement des valeurs aberrantes : `Duree_Traitement_Sec` (< 10s = `null`) et `Prix_Prestation` (< 12 € = `null`).

###  Table Prestataire D
1. Traduction globale des libellés et valeurs de l'anglais vers le français (ex: `Social Media` → `Réseaux_Sociaux`).
2. Standardisation des identifiants agents (`AG101` → `AG001`).
3. Correction des anomalies de formats de date inversés (`MM/JJ/AAAA` → `JJ/MM/AAAA`).
4. Traitement des valeurs aberrantes sur les durées et conversion des booléens en `Oui` / `Non`.

---

##  4. Volet Excel : Contrôle, TCD & Macros VBA

En amont de Power BI, un outil opérationnel structuré sous Excel a été mis en place autour de deux pages :

###  Page 1 : « Contrôle de KPI »
* **Calculs avancés :** Utilisation de formules complexes (`SOMME.SI.ENS`, `NB.SI.ENS`, etc.) pour agréger dynamiquement les indicateurs clés.
* **Mise en forme conditionnelle :** Alertes visuelles instantanées en cas de dépassement de seuil ou de non-conformité.
* **Macros VBA :**
  * Macro de **basculement d'affichage** pour activer/désactiver la mise en forme conditionnelle.
  * Macro d'**exportation automatique** de la vue sous format PDF.

### 📊 Page 2 : « Dashboard Excel »
* **TCD multi-angles :** Tableaux croisés dynamiques structurés par familles d'indicateurs (Volume, Qualité, Productivité, Coûts).
* **Restitution visuelle :** Graphiques directement connectés aux TCD pour un pilotage synthétique.
* **Reporting journalier :** Macro d'**exportation PDF en un clic** pour la génération du rapport quotidien.

---

##  5. Modélisation des Données (Schéma en Étoile)

L'ensemble des tables nettoyées a été combiné (`Append`) dans une table de faits unique reliée à ses dimensions selon un schéma en étoile (`1:N`) à filtrage unidirectionnel :

* **Table de Faits :**
  * `Appels_Global` : Contient l'historique exhaustif des interactions.
* **Tables de Dimensions :**
  * `Dim_Agent` (`Id Agent`, `Nombre d'heures travaillées`, `Nombre d'appels traités`, `Prestataire`)
  * `Dim_Prestataire` (`Id Prestataire`, `Nom Prestataire`)
  * `Dim_Date` (`Id Mois`, `Mois`, `Date_Appel`)

---

##  6. Dashboard Power BI & Familles de KPI

Le tableau de bord Power BI est structuré en 6 pages interactives permettant d'analyser la performance sous plusieurs angles :

### 1.  Vue d'Ensemble
* **Volume global :** ~9,58K appels reçus | Montant total facturé : 279 k€
* **Taux de résolution :** 58,78 % | **Satisfaction moyenne :** 3,78 / 5
* **Backlog :** 2,9K appels en attente
* **Répartition canaux :** Téléphone (43,7%), Courrier (21,1%), Email (20,9%), Chat, Réseaux Sociaux.

### 2.  Contrôle Qualité
* **Taux de conformité global :** 76,75 %
* **Taux d'erreur :** 23,25 % | **Taux de réclamation :** 21,19 %
* **Analyse comparative :** Suivi des volumes conformes vs non-conformes et taux de reprise par prestataire.

### 3.  Délais & Temps de Traitement
* **Durée Moyenne de Traitement (DMT) :** 266,5 secondes (Min : 10s | Max : 684s).
* **Volume total de traitement :** ~697 heures.
* **Analyses croisées :** DMT par motif, par canal et par statut du ticket.

### 4.  Productivité & Performance des Agents
* **Agents actifs :** 34 agents
* **Taux de productivité global :** 69,35 % | **Taux d'occupation :** 58,54 %
* **Analyse individuelle :** Vue détaillée par agent (heures travaillées, volume traité, taux de résolution, taux de reprise).

### 5.  Satisfaction Client (NPS & CSAT)
* **Score NPS moyen :** 6,92 / 10
* **Corrélation :** Étude de l'impact du DMT sur les notes NPS et CSAT.
* **Classement prestataires :** Comparatif de la satisfaction client selon le prestataire.

### 6.  Facturation & Suivi Budgétaire
* **Budget facturé :** 279 184 € (Objectif : 300 000 €).
* **Coût moyen par appel :** ~31 €
* **Ventilation des coûts :** Par prestataire, par canal et par motif.

---

##  7. Outils & Technologies Utilises

* **Microsoft Excel & Power Query :** Nettoyage des données, formules complexes (`SOMME.SI.ENS`, `NB.SI.ENS`, `INDEX + EQUIV`, `RECHERCHEV/X` etc), TCD et visualisation.
* **VBA (Visual Basic for Applications) :** Automatisations (exports PDF, mises en forme conditionnel).
* **Power BI Desktop :** Modélisation relationnelle (Star Schema), mesures DAX avancées, visualisations et filtres dynamiques.

---

##  8. Structure du Repository

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
