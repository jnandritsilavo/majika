VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'    888 888~~       e      888b    |       888b    |      e      888~-_   888 Y88b      / 888~~  888       ,88~-_
'    888 888___     d8b     |Y88b   |       |Y88b   |     d8b     888   \  888  Y88b    /  888___ 888      d888   \
'    888 888       /Y88b    | Y88b  |       | Y88b  |    /Y88b    888    | 888   Y88b  /   888    888     88888    |
'    888 888      /  Y88b   |  Y88b |       |  Y88b |   /  Y88b   888   /  888    Y888/    888    888     88888    |
'|   88P 888     /____Y88b  |   Y88b|       |   Y88b|  /____Y88b  888_-~   888     Y8/     888    888      Y888   /
' \__8"  888___ /      Y88b |    Y888       |    Y888 /      Y88b 888 ~-_  888      Y      888___ 888____   `88_-~
                                                                                                                    
'888~~\    ,88~-_   888  /   888~~  ~~~888~~~ 888~-_        e            888     ,e,
'888   |  d888   \  888 /    888___    888    888   \      d8b           888      "    /~~~8e  888-~88e   /~~~8e
'888 _/  88888    | 888/\    888       888    888    |    /Y88b          888     888       88b 888  888       88b
'888  \  88888    | 888  \   888       888    888   /    /  Y88b         888     888  e88~-888 888  888  e88~-888
'888   |  Y888   /  888   \  888       888    888_-~    /____Y88b        888     888 C888  888 888  888 C888  888
'888__/    `88_-~   888    \ 888___    888    888 ~-_  /      Y88b       888____ 888  "88_-888 888  888  "88_-888


' Dépot GitHub : https://github.com/jnandritsilavo/majika
                                                                                                                    
Sub Enregistrer_Word()
    ' Déclaration des variables
    Dim iR As Integer ' Nombre total d'enregistrements
    Dim i As Integer ' Compteur pour la boucle
    Dim oDoc As Document ' Référence au document actif
    Dim DocName As String ' Nom de fichier du document é enregistrer
    Dim oDS As MailMergeDataSource ' Source de données pour le publipostage
    Dim outputPath As String ' Chemin de sauvegarde des fichiers générés

    ' Affectation des objets
    Set oDoc = ActiveDocument ' Récupére le document actif
    Set oDS = oDoc.MailMerge.DataSource ' Récupére la source de données de publipostage

    ' Récupération du chemin du document actif pour créer un dossier de destination
    outputPath = oDoc.Path & "\MAJIKA_WORD\" ' Dossier de sauvegarde pour les fichiers Word

    ' Récupération du nombre d'enregistrements dans la source de données
    iR = oDoc.MailMerge.DataSource.RecordCount
    Debug.Print iR ' Affiche le nombre total d'enregistrements dans la fenétre de débogage

    ' Boucle pour traiter chaque enregistrement dans la source de données
    For i = 1 To iR
        With oDoc.MailMerge
            ' Définition du premier et dernier enregistrement pour chaque itération
            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            ' Crée un nouveau document en utilisant l'enregistrement actuel
            .Destination = wdSendToNewDocument
            ' Exécution du publipostage
            .Execute
            ' Active l'enregistrement en cours pour obtenir les valeurs des champs
            .DataSource.ActiveRecord = i
            ' Concaténation de deux champs pour créer le nom du document
            DocName = .DataSource.DataFields(1).Value ' Premier champ pour le nom du document
            DocName = DocName & "_" & .DataSource.DataFields(2).Value ' Ajout du deuxiéme champ pour compléter le nom
            Debug.Print DocName; i ' Affiche le nom du document et l'index actuel dans la fenétre de débogage
        End With

        ' Sauvegarde du document généré sous forme de fichier Word
        With ActiveDocument
            .SaveAs outputPath & DocName & ".doc" ' Enregistre le fichier dans le dossier spécifié
            .Close ' Ferme le document aprés l'enregistrement
        End With
    Next i
    ' Message pour indiquer la fin du traitement
    MsgBox "Le traitement est terminé. Tous les documents ont été enregistrés.", vbInformation
End Sub

Sub Enregistrer_Pdf()
    ' Déclaration des variables
    Dim iR As Integer ' Nombre total d'enregistrements
    Dim i As Integer ' Compteur pour la boucle
    Dim oDoc As Document ' Référence au document actif
    Dim DocName As String ' Nom de fichier du document é enregistrer
    Dim oDS As MailMergeDataSource ' Source de données pour le publipostage
    Dim outputPath As String ' Chemin de sauvegarde des fichiers PDF générés

    ' Affectation des objets
    Set oDoc = ActiveDocument ' Récupére le document actif
    Set oDS = oDoc.MailMerge.DataSource ' Récupére la source de données de publipostage

    ' Récupération du chemin du document actif pour créer un dossier de destination
    outputPath = oDoc.Path & "\MAJIKA_PDF\" ' Dossier de sauvegarde pour les fichiers PDF

    ' Récupération du nombre d'enregistrements dans la source de données
    iR = oDoc.MailMerge.DataSource.RecordCount
    Debug.Print iR ' Affiche le nombre total d'enregistrements dans la fenétre de débogage

    ' Boucle pour traiter chaque enregistrement dans la source de données
    For i = 1 To iR
        With oDoc.MailMerge
            ' Définition du premier et dernier enregistrement pour chaque itération
            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            ' Crée un nouveau document en utilisant l'enregistrement actuel
            .Destination = wdSendToNewDocument
            ' Exécution du publipostage
            .Execute
            ' Active l'enregistrement en cours pour obtenir les valeurs des champs
            .DataSource.ActiveRecord = i
            ' Concaténation de deux champs pour créer le nom du document
            DocName = .DataSource.DataFields(1).Value ' Premier champ pour le nom du document
            DocName = DocName & "_" & .DataSource.DataFields(2).Value ' Ajout du deuxiéme champ pour compléter le nom
            Debug.Print DocName; i ' Affiche le nom du document et l'index actuel dans la fenétre de débogage
        End With

        ' Sauvegarde du document généré au format PDF
        With ActiveDocument
            .ExportAsFixedFormat outputPath & DocName & ".pdf", _
                ExportFormat:=wdExportFormatPDF
            .Close False ' Ne pas enregistrer les modifications
        End With
    Next i
    ' Message pour indiquer la fin du traitement
    MsgBox "Le traitement est terminé. Tous les documents ont été enregistrés.", vbInformation
End Sub

