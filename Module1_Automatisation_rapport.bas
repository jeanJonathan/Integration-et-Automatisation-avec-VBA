Attribute VB_Name = "Module1"
'Connexion a la bd et affichage des datas
'On veut creer une macro VBA qui se connecte � votre base de donn�es, ex�cute les requ�tes SQL n�cessaires et importe les donn�es dans Excel.
Sub extractionDonnees()

    Dim conn As Object, rs As Object, strSql1 As String, strSql2 As String
    ' conn : Variable pour l'objet de connexion � la base de donn�es.
    ' rs : Variable pour l'objet Recordset qui contiendra les r�sultats de la requ�te SQL.
    ' strSql1, strSql2 : Variables pour stocker les requ�tes SQL sous forme de cha�nes de caract�res.
    
    ' Cr�ation de la Connexion et Ouverture de la connexion
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=MysqlTestSource"
    ' Set conn = CreateObject("ADODB.Connection") : Cr�e une nouvelle instance d'un objet de connexion ADODB.
    ' conn.Open "DSN=MysqlTestSource" : Ouvre la connexion � la base de donn�es MySQL en utilisant le DSN (Data Source Name) configur� sous le nom "MysqlTestSource".
    
    ' Requ�tes SQL
    strSql1 = "SELECT date_emission, montant, etat, description FROM factures;"
    strSql2 = "SELECT date_paiement, montant, mode_paiement FROM paiements;"
    strSQL3 = "SELECT date_intervention, type_intervention, cout, temps_passe, cout_pieces FROM interventions;"
    
    ' Cr�ation d'une instance d'un objet Recordset
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Ex�cute la requ�te SQL contenue dans strSql1 et place les r�sultats dans l'objet rs
    Set rs = conn.Execute(strSql1)
    
    ' Copie des datas
    ThisWorkbook.Worksheets("Feuil1").Range("B2").CopyFromRecordset rs
    ' ThisWorkbook permet d'acc�der directement � ce classeur Excel sans avoir besoin de sp�cifier son nom.
    ' Chaque objet Worksheet repr�sente une feuille de calcul individuelle dans le classeur.
    'Range pour sp�cifier la cellule en question dans la feuille de calcul

    Set rs = conn.Execute(strSql2)
    ThisWorkbook.Worksheets("Feuil1").Range("F2").CopyFromRecordset rs

    Set rs = conn.Execute(strSQL3)
    ThisWorkbook.Worksheets("Feuil1").Range("J2").CopyFromRecordset rs
    
    ' Fermeture
    rs.Close ' Ferme l'objet Recordset.
    conn.Close ' Ferme la connexion � la base de donn�es.
    Set rs = Nothing ' Lib�re l'objet Recordset de la m�moire.
    Set conn = Nothing ' Lib�re l'objet de connexion de la m�moire.
    
End Sub

'Automatisation des calculs des moyennes et les totaux des datas
'Une procedure pour calculer les totaux et les moyennes n�cessaires pour les rapports.
Sub calculTotauxEtMoy()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Feuil1")

    'Montant totals des factures et moyennes
    ws.Range("D22").Value = "Total Montant factures"
    'Formula est la methode qui permet de d�finir une formule Excel.
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
'On veut formater les donn�es pour qu'elles soient plus lisibles et pr�sentables.
Sub formatageRapport()
    ' D�clare une variable pour la feuille de calcul
    Dim ws As Worksheet
    ' Assigne la feuille de calcul "Feuil1" � la variable ws
    Set ws = ThisWorkbook.Sheets("Feuil1")
    
    ' D�finition des titres pour les sections
    ws.Range("B1:E1").Merge ' Fusionne les cellules de B1 � E1 pour cr�er un seul titre
    ws.Range("B1").Value = "Factures" ' D�finit le titre pour les factures
    ws.Range("B1").Font.Bold = True ' Met le titre en gras
    ws.Range("B1").Interior.Color = RGB(169, 208, 142) ' Applique une couleur de fond verte claire
    
    ws.Range("F1:H1").Merge ' Fusionne les cellules de F1 � H1 pour cr�er un seul titre
    ws.Range("F1").Value = "Paiements" ' D�finit le titre pour les paiements
    ws.Range("F1").Font.Bold = True ' Met le titre en gras
    ws.Range("F1").Interior.Color = RGB(142, 169, 219) ' Applique une couleur de fond bleue claire
    
    ws.Range("J1:N1").Merge ' Fusionne les cellules de J1 � N1 pour cr�er un seul titre
    ws.Range("J1").Value = "Interventions" ' D�finit le titre pour les interventions
    ws.Range("J1").Font.Bold = True ' Met le titre en gras
    ws.Range("J1").Interior.Color = RGB(255, 192, 0) ' Applique une couleur de fond orange claire
    
    ' Formatage des en-t�tes de colonnes
    ws.Range("B2:E2").Font.Bold = True ' Met les en-t�tes de colonnes en gras pour la section Factures
    ws.Range("F2:H2").Font.Bold = True ' Met les en-t�tes de colonnes en gras pour la section Paiements
    ws.Range("J2:N2").Font.Bold = True ' Met les en-t�tes de colonnes en gras pour la section Interventions
    
    ' Application des bordures
    With ws.Range("B1:E21").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous ' Ligne continue
        .ColorIndex = 0 ' Couleur noire
        .TintAndShade = 0 ' Pas de nuance
        .Weight = xlThin ' �paisseur fine
    End With
    
    With ws.Range("F1:H21").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous ' Ligne continue
        .ColorIndex = 0 ' Couleur noire
        .TintAndShade = 0 ' Pas de nuance
        .Weight = xlThin ' �paisseur fine
    End With
    
    With ws.Range("J1:N21").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous ' Ligne continue
        .ColorIndex = 0 ' Couleur noire
        .TintAndShade = 0 ' Pas de nuance
        .Weight = xlThin ' �paisseur fine
    End With
    
    ' Formatage des totaux et moyennes
    ws.Range("C22:D23").Font.Bold = True ' Met les totaux et moyennes en gras
    ws.Range("C22:D23").Interior.Color = RGB(255, 255, 153) ' Applique une couleur de fond jaune claire
    
    ws.Range("G22:H23").Font.Bold = True ' Met les totaux et moyennes en gras
    ws.Range("G22:H23").Interior.Color = RGB(255, 255, 153) ' Applique une couleur de fond jaune claire
    
    ws.Range("L22:O23").Font.Bold = True ' Met les totaux et moyennes en gras
    ws.Range("L22:O23").Interior.Color = RGB(255, 255, 153) ' Applique une couleur de fond jaune claire
    
    ' Ajustement de la largeur des colonnes pour une meilleure lisibilit�
    ws.Columns("B:N").AutoFit ' Ajuste automatiquement la largeur des colonnes pour s'adapter au contenu
End Sub

'Automatisation de l'Envoi d'Emails
'On veut cr�er une macro pour envoyer automatiquement les rapports par email en utilisant Outlook.
' Automatisation de l'Envoi d'Emails
' On veut cr�er une macro pour envoyer automatiquement les rapports par email en utilisant Outlook.
Sub EnvoieEmailavecRapport()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Feuil1")

    ' Cr�er une instance d'Outlook
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error GoTo 0

    If OutApp Is Nothing Then
        MsgBox "Outlook n'est pas install� ou configur� correctement sur votre ordinateur.", vbExclamation
        Exit Sub
    End If

    ' D�finir le contenu de l'email
    On Error Resume Next
    With OutMail
        .To = "jean-jonathan.koffi@etud.univ-pau.fr" ' Adresse email du destinataire
        .CC = "" ' Adresse(s) en copie
        .BCC = "" ' Adresse(s) en copie cach�e
        .Subject = "Rapport de Maintenance et Facturation" ' Sujet de l'email
        .Body = "Veuillez trouver ci-joint le rapport de maintenance et facturation." ' Corps de l'email
        .Attachments.Add ThisWorkbook.FullName ' Attacher le fichier Excel actuel
        .Send ' Envoyer l'email
    End With
    On Error GoTo 0

    If Err.Number <> 0 Then
        MsgBox "Une erreur s'est produite lors de l'envoi de l'email. Veuillez v�rifier votre configuration Outlook.", vbExclamation
    Else
        MsgBox "Email envoy� avec succ�s.", vbInformation
    End If

    ' Lib�rer les objets
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

