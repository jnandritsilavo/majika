
# Publipostage Automatisé - Word et PDF

Ce programme VBA permet d'effectuer un publipostage automatisé à partir de fichiers Excel et Word, générant des documents Word et PDF pour chaque enregistrement dans la source de données. Il utilise un modèle de document Word (`MODELY`) et une source de données Excel (`LISITRA`).

## Prérequis

Pour que le programme fonctionne, assurez-vous de disposer des éléments suivants :
1. Un fichier Word nommé `MODELY`, qui servira de modèle de document pour le publipostage.
2. Un fichier Excel nommé `LISITRA`, qui contiendra les données pour le publipostage.
3. Deux dossiers :
   - `MAJIKA_PDF` pour stocker les fichiers PDF générés.
   - `MAJIKA_WORD` pour stocker les fichiers Word générés.

## Compétences Requises

Pour manipuler le programme et adapter les fonctionnalités selon les besoins, il est recommandé d'avoir les compétences suivantes :

1. **Microsoft Word et Excel** :
   - Compréhension du publipostage dans Word, notamment l'importation de données et l'utilisation de modèles.
   - Organisation des données dans Excel pour servir de source de données fiable.

2. **Programmation en VBA (Visual Basic for Applications)** :
   - Connaissances de base en macros, notamment leur création, exécution, et modification.
   - Compétences dans l’utilisation des boucles, des variables et de la manipulation des objets comme `Document` et `MailMergeDataSource`.
   - Familiarité avec les méthodes de gestion de fichiers en VBA pour enregistrer et manipuler des documents (Word, PDF).

3. **Organisation de fichiers et dossiers** :
   - Capacité à organiser les dossiers et fichiers (dossiers `MAJIKA_PDF` et `MAJIKA_WORD`) pour stocker les fichiers générés.

4. **Débogage VBA** :
   - Compétence dans l'utilisation de la fenêtre de débogage (`Debug.Print`) pour vérifier les valeurs des variables et résoudre les erreurs éventuelles.

5. **Automatisation de tâches** :
   - Savoir exploiter les fonctionnalités de Microsoft Office pour automatiser des tâches répétitives, afin de gagner du temps et minimiser les erreurs humaines.

### Compétences recommandées pour un usage avancé
- **Gestion de chemins dynamiques** : Pouvoir adapter le code pour des chemins de sauvegarde variés.
- **Gestion d'erreurs en VBA** : Connaître les techniques de gestion des erreurs pour éviter les interruptions de script.

## Fonctionnalités

Le programme est constitué de deux sous-programmes :

- **Enregistrer_Word** : génère un document Word pour chaque enregistrement de la source de données et le sauvegarde dans le dossier `MAJIKA_WORD`.
- **Enregistrer_Pdf** : génère un document Word pour chaque enregistrement et l'enregistre au format PDF dans le dossier `MAJIKA_PDF`.

## Utilisation

1. Placez les fichiers `MODELY.docx` et `LISITRA.xlsx` dans le même répertoire que ce code.
2. Créez les dossiers `MAJIKA_WORD` et `MAJIKA_PDF` dans le même répertoire.
3. Exécutez les macros `Enregistrer_Word` et `Enregistrer_Pdf` selon le type de fichier que vous souhaitez générer.

## Code VBA

Le code utilise des champs de données dans `LISITRA` pour nommer les documents générés. Assurez-vous que `LISITRA` contient au moins deux colonnes pour que le nom de fichier se compose des données de ces champs.

### Exemple de structure des macros

```vba
Sub Enregistrer_Word()
    ' Déclaration des variables
    ' Code pour générer et enregistrer les fichiers Word
End Sub

Sub Enregistrer_Pdf()
    ' Déclaration des variables
    ' Code pour générer et enregistrer les fichiers PDF
End Sub
```

## Notes

- Le programme vérifie automatiquement le nombre d'enregistrements dans la source de données Excel.
- Après la génération des documents, un message indique la fin du traitement.

## Auteurs

JEAN NARIVELO
