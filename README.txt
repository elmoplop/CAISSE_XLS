Version 1.2
Ce programme permet de lire un rapport html issu du site https://caisse.enregistreuse.fr/ et de créer plusieur fichier xlsx à partir des données extraites.

Pour l'utiliser:
- Extraire le rapport de caisse souhaite à partir du l'onglet rapport du site https://caisse.enregistreuse.fr/
- Le placer au même endroit que CAISSE_XSLX.exe
- Lancer CAISSE_XSLX.exe

Les tables resultantes sont crées dans un sous dossier GEN_[DATE]

Changelog :
Version 1.2 - 18/12/21 
- Suppression des suffixe aux fichiers
- Les fichiers sont maintenant indexé à la date de couverture plutot que la date d'execution

Version 1.1 - 08/12/21 
- BUGFIX - Le dernier magasin de la liste n'était pris en considération pour la création de l'onglet, du fichier in dependant et de l'indicatif d'erreur
- Ajout du calcul de la facturation a partir du % de commission
- Mise en forme des cellules monétaires