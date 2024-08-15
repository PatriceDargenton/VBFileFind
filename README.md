# VBFileFind
Recherche de fichiers pour remplacer celle de Windows

## Table des matières
- [Utilisation](#utilisation)
- [Limitations](#limitations)
- [Projets](#projets)
- [Versions](#versions)

## Utilisation
Lancer VBFileFind en mode administrateur, ajouter le menu contextuel via le bouton +, quitter VBFileFind, puis lancer une recherche via l'explorateur de fichier de Windows, en sélectionnant un dossier avec le bouton droit de la souris, puis en cliquant sur "Rechercher avec VBFileFind" (avec Windows 11, il faudra appuyer sur la touche majuscule pour afficher directement ce menu, sinon il faudra d'abord afficher le menu Autres options).

## Limitations
- Windows 11 : la recherche d'un mot dans le bloc-notes ne peut plus être pilotée via les envois de touche : désactivé
- Calcul des tailles de dossier : ce n'est pas fiable à 100%
- Recherche UTF8 : en .NET la représentation des chaînes utilise par défaut l'UTF-16 : chaque caractère est représenté par 2 octets (ou 4 octets pour les caractères en dehors du Basic Multilingual Plane). Lorsque les fichiers sont encodés en UTF8, la taille de l'encodage peut varier, chaque caractère simple est encodé sur un octet, et les autres caractères sur 2 ou plus de caractères, il faut procéder autrement (on ne peut donc pas chercher des mots avec un accent, essayez avec le fichier limitation.txt)

## Projets

## Versions

Voir le [Changelog.md](Changelog.md)

Documentation d'origine complète : [VBFileFind.html](http://patrice.dargenton.free.fr/CodesSources/VBFileFind.html)