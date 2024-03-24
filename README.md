Nom du projet : Auto_Agenda 

Objectif : Automatisation du formatage du calendrier des cours de master SIGMA de format Excel vers le format ical/ics

Produit développé : Un code réalisé en langage Python qui permet de créer et synchroniser un Agenda Google contenant le calandrier des cours de master SIGMA avec la possibilité de faire une sélection par intervenant ou par UE
Démarche du travail: 

* Etape 1 : formatage de base des données intiale, qui réprésente un fichier Excel contenant les calandriers des cours de Master SIGMA pour l'année 2023/2024, et création d'un nouveau fichier de la même format .xls mais plus structuré en excutant un script python

  
* Etape 2 : Publication de l’agenda en ligne : Pour réaliser cette tâche, nous avons choisi d’utiliser la plateforme Google Calendar qui supporte le format ics/ical.
  
Mais avant de réaliser cette étape, une démarche de configuration de l’environnement de développement (via l’activation d’API Google et l’obtention des identifiants et le fichier .json) sera nécessaire.


Création d'un code Python qui facilitera l’exportation  du contenu du fichier Excel précédemment créé vers l'Agenda Google. 

Enfin, pour synchroniser les changements effectués sur le fichier Excel initial, il suffit de sauvegarder ces changements, puis de relancer le code pour actualiser l’agenda publié sur internet.


* Etape 3 : Exportation du calandrier pour retourner à la base de données initiale  via l'excution d'un deuxiéme script python 
  

