Attribute VB_Name = "Module1"
'Connexion a la bd et affichage des datas
'On veut creer une macro VBA qui se connecte à votre base de données, exécute les requêtes SQL nécessaires et importe les données dans Excel.
Sub extractionDonnees()

    Dim conn As Object, rs As Object, strSql1 As String, strSql2 As String
    ' conn : Variable pour l'objet de connexion à la base de données.
    ' rs : Variable pour l'objet Recordset qui contiendra les résultats de la requête SQL.
    ' strSql1, strSql2 : Variables pour stocker les requêtes SQL sous forme de chaînes de caractères.
    
    ' Création de la Connexion et Ouverture de la connexion
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=MysqlTestSource"
    ' Set conn = CreateObject("ADODB.Connection") : Crée une nouvelle instance d'un objet de connexion ADODB.
    ' conn.Open "DSN=MysqlTestSource" : Ouvre la connexion à la base de données MySQL en utilisant le DSN (Data Source Name) configuré sous le nom "MysqlTestSource".
    
    ' Requêtes SQL
    strSql1 = "SELECT date_emission, montant, etat, description FROM factures;"
    strSql2 = "SELECT date_paiement, montant, mode_paiement FROM paiements;"
    strSQL3 = "SELECT date_intervention, type_intervention, cout, temps_passe, cout_pieces FROM interventions;"
    
    ' Création d'une instance d'un objet Recordset
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Exécute la requête SQL contenue dans strSql1 et place les résultats dans l'objet rs
    Set rs = conn.Execute(strSql1)
    
    ' Copie des datas
    ThisWorkbook.Worksheets("Feuil1").Range("B2").CopyFromRecordset rs
    ' ThisWorkbook permet d'accéder directement à ce classeur Excel sans avoir besoin de spécifier son nom.
    ' Chaque objet Worksheet représente une feuille de calcul individuelle dans le classeur.
    'Range pour spécifier la cellule en question dans la feuille de calcul

    Set rs = conn.Execute(strSql2)
    ThisWorkbook.Worksheets("Feuil1").Range("F2").CopyFromRecordset rs

    Set rs = conn.Execute(strSQL3)
    ThisWorkbook.Worksheets("Feuil1").Range("J2").CopyFromRecordset rs
    
    ' Fermeture
    rs.Close ' Ferme l'objet Recordset.
    conn.Close ' Ferme la connexion à la base de données.
    Set rs = Nothing ' Libère l'objet Recordset de la mémoire.
    Set conn = Nothing ' Libère l'objet de connexion de la mémoire.
    
End Sub

'Automatisation des calculs des moyennes et les totaux des datas
'Une procedure pour calculer les totaux et les moyennes nécessaires pour les rapports.
Sub calculTotauxEtMoy()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Feuil1")

    'Montant totals des factures et moyennes
    ws.Range("D22").Value = "Total Montant factures"
    'Formula est la methode qui permet de définir une formule Excel.
    'ws.Range("D22").Formula = "=SUM(D2:D" & ws.Cells(ws.Rows.Count, "C").End(xlUp).Row & ")"
    ws.Range("C22").Value = WorksheetFunction.Sum(ws.Range(ws.Range("C2"), ws.Range("C21")))
    
    'Moyenne des factures
    ws.Range("D23").Value = "Moyenne factures"
    'ws.Range("D23").Formula = "=AVERAGE(D2:D" & ws.Cells(ws.Rows.Count, "C").End(xlUp).Row & ")"
    ws.Range("C23").Value = WorksheetFunction.Average(ws.Range(ws.Range("C2"), ws.Range("C21")))
    
    'Total paiements et moyennes
    ws.Range("H22").Value = "Total paiements"
    ws.Range("H23").Value = "Moyenne paiements"
    
    ws.Range("G22").Value = WorksheetFunction.Sum(ws.Range(ws.Range("G2"), ws.Range("G21")))
    ws.Range("G23").Value = WorksheetFunction.Average(ws.Range(ws.Range("G2"), ws.Range("G21")))
    
    'Total cout et moyennes
    ws.Range("O22").Value = "Total cout"
    ws.Range("O23").Value = "Moyenne cout"
    
    'Total cout et moyennes
    ws.Range("L22").Value = WorksheetFunction.Sum(ws.Range(ws.Range("L2"), ws.Range("L21")))
    ws.Range("L23").Value = WorksheetFunction.Average(ws.Range(ws.Range("L2"), ws.Range("L21")))
    
End Sub


'Mise en forme des datas
'On veut formater les données pour qu'elles soient plus lisibles et présentables.
Sub formatageRapport()
    ' Déclare une variable pour la feuille de calcul
    Dim ws As Worksheet
    ' Assigne la feuille de calcul "Feuil1" à la variable ws
    Set ws = ThisWorkbook.Sheets("Feuil1")
    
    ' Définition des titres pour les sections
    ws.Range("B1:E1").Merge ' Fusionne les cellules de B1 à E1 pour créer un seul titre
    ws.Range("B1").Value = "Factures" ' Définit le titre pour les factures
    ws.Range("B1").Font.Bold = True ' Met le titre en gras
    ws.Range("B1").Interior.Color = RGB(169, 208, 142) ' Applique une couleur de fond verte claire
    
    ws.Range("F1:H1").Merge ' Fusionne les cellules de F1 à H1 pour créer un seul titre
    ws.Range("F1").Value = "Paiements" ' Définit le titre pour les paiements
    ws.Range("F1").Font.Bold = True ' Met le titre en gras
    ws.Range("F1").Interior.Color = RGB(142, 169, 219) ' Applique une couleur de fond bleue claire
    
    ws.Range("J1:N1").Merge ' Fusionne les cellules de J1 à N1 pour créer un seul titre
    ws.Range("J1").Value = "Interventions" ' Définit le titre pour les interventions
    ws.Range("J1").Font.Bold = True ' Met le titre en gras
    ws.Range("J1").Interior.Color = RGB(255, 192, 0) ' Applique une couleur de fond orange claire
    
    ' Formatage des en-têtes de colonnes
    ws.Range("B2:E2").Font.Bold = True ' Met les en-têtes de colonnes en gras pour la section Factures
    ws.Range("F2:H2").Font.Bold = True ' Met les en-têtes de colonnes en gras pour la section Paiements
    ws.Range("J2:N2").Font.Bold = True ' Met les en-têtes de colonnes en gras pour la section Interventions
    
    ' Application des bordures
    With ws.Range("B1:E21").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous ' Ligne continue
        .ColorIndex = 0 ' Couleur noire
        .TintAndShade = 0 ' Pas de nuance
        .Weight = xlThin ' Épaisseur fine
    End With
    
    With ws.Range("F1:H21").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous ' Ligne continue
        .ColorIndex = 0 ' Couleur noire
        .TintAndShade = 0 ' Pas de nuance
        .Weight = xlThin ' Épaisseur fine
    End With
    
    With ws.Range("J1:N21").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous ' Ligne continue
        .ColorIndex = 0 ' Couleur noire
        .TintAndShade = 0 ' Pas de nuance
        .Weight = xlThin ' Épaisseur fine
    End With
    
    ' Formatage des totaux et moyennes
    ws.Range("C22:D23").Font.Bold = True ' Met les totaux et moyennes en gras
    ws.Range("C22:D23").Interior.Color = RGB(255, 255, 153) ' Applique une couleur de fond jaune claire
    
    ws.Range("G22:H23").Font.Bold = True ' Met les totaux et moyennes en gras
    ws.Range("G22:H23").Interior.Color = RGB(255, 255, 153) ' Applique une couleur de fond jaune claire
    
    ws.Range("L22:O23").Font.Bold = True ' Met les totaux et moyennes en gras
    ws.Range("L22:O23").Interior.Color = RGB(255, 255, 153) ' Applique une couleur de fond jaune claire
    
    ' Ajustement de la largeur des colonnes pour une meilleure lisibilité
    ws.Columns("B:N").AutoFit ' Ajuste automatiquement la largeur des colonnes pour s'adapter au contenu
End Sub

'Automatisation de l'Envoi d'Emails
'On veut créer une macro pour envoyer automatiquement les rapports par email en utilisant Outlook.
' Automatisation de l'Envoi d'Emails
' On veut créer une macro pour envoyer automatiquement les rapports par email en utilisant Outlook.
Sub EnvoieEmailavecRapport()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Feuil1")

    ' Créer une instance d'Outlook
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error GoTo 0

    If OutApp Is Nothing Then
        MsgBox "Outlook n'est pas installé ou configuré correctement sur votre ordinateur.", vbExclamation
        Exit Sub
    End If

    ' Définir le contenu de l'email
    On Error Resume Next
    With OutMail
        .To = "jean-jonathan.koffi@etud.univ-pau.fr" ' Adresse email du destinataire
        .CC = "" ' Adresse(s) en copie
        .BCC = "" ' Adresse(s) en copie cachée
        .Subject = "Rapport de Maintenance et Facturation" ' Sujet de l'email
        .Body = "Veuillez trouver ci-joint le rapport de maintenance et facturation." ' Corps de l'email
        .Attachments.Add ThisWorkbook.FullName ' Attacher le fichier Excel actuel
        .Send ' Envoyer l'email
    End With
    On Error GoTo 0

    If Err.Number <> 0 Then
        MsgBox "Une erreur s'est produite lors de l'envoi de l'email. Veuillez vérifier votre configuration Outlook.", vbExclamation
    Else
        MsgBox "Email envoyé avec succès.", vbInformation
    End If

    ' Libérer les objets
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub



'main
Sub main()
    extractionDonnees
    calculTotauxEtMoy
    formatageRapport
    EnvoieEmailavecRapport
End Sub

