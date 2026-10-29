# 🎓 Pipeline ETL & Progiciel Data - École Doctorale DESPEG

> ⚠️ **Note sur la confidentialité et la propriété intellectuelle**
> Ce projet ayant été développé dans le cadre de ma mission pour l'Université Côte d'Azur (ED DESPEG), le code source complet du progiciel et les bases de données associées relèvent de la propriété interne de l'établissement.
> Pour respecter la confidentialité des données et l'éthique professionnelle :
> * Les données présentées dans les rapports Power BI de ce portfolio sont **100 % fictives**. Elles ont été générées de toutes pièces pour illustrer le fonctionnement des outils sans exposer d'informations réelles.
> * Seuls des **extraits de code ciblés** (snippets) sont partagés ici afin de démontrer la logique algorithmique (ETL, gestion des erreurs) et la structure de l'interface graphique.

## 📝 Contexte du Projet
Conception et déploiement de A à Z d'un progiciel métier pour automatiser le suivi des effectifs et l'analyse de l'insertion professionnelle des jeunes docteurs. L'outil remplace un traitement manuel chronophage par un pipeline de données robuste, utilisable de manière autonome par l'équipe administrative.

## 🛠️ Stack Technique
* **Data Engineering (ETL) :** Python (Pandas, Numpy, Regex) pour le nettoyage textuel (NFKD), la détection de doublons et le croisement de bases.
* **Interface & Déploiement :** CustomTkinter (GUI) et environnement Python portable pour contourner les restrictions d'installation et d'antivirus.
* **Contrôle Qualité :** XlsxWriter (génération de fichiers d'audit avec menus déroulants en cascade et mise en forme conditionnelle).
* **Business Intelligence :** Microsoft Power BI, requêtes DAX, génération de KPIs.

## 📂 Contenu de ce répertoire

* **`etl_architecture_showcase.py`** : Extraits de code (Snippets) illustrant l'architecture du projet (gestion des crashs via *sys.excepthook*, conception de l'interface graphique, nettoyage de données, automatisation Excel).
* **`Faux_Dashboard_Stat_ED_DESPEG.pdf`** & **`Faux_Dashboard_insertion_pro.pdf`** : Restitution visuelle des indicateurs institutionnels générée à partir de la Mock Data.
* **`GUIDE_UTILISATEUR.pdf`** : Documentation métier rédigée pour la prise en main de l'outil par l'équipe administrative.
* **`Référentiel_Technique_de_l_Outils.pdf`** : Documentation technique de l'architecture logicielle à destination des développeurs.
* **`Rapport_de_stage_Thibault_PHILIBERT.pdf`** : Rapport complet incluant le détail du projet d'ingénierie et une étude économétrique (modèles Logit/MCO) sur les déterminants de l'achèvement de la thèse.
