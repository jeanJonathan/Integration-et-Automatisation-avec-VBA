VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Formulaire de facturation"
   ClientHeight    =   6468
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6576
   OleObjectBlob   =   "Formulaire de facturation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bEnregistrer_Click()

    Dim nomClient As String
    Dim adresseClient As String
    Dim contactClient As String
    Dim emailClient As String
    Dim dateEmission As Date
    Dim modalitePaiement As String
    Dim delaisPaiement As Integer
    Dim descriptionProduit As String
    Dim quantiteProduit As Integer
    Dim tauxTVA As Double
    Dim prixUnitaire As Double
    Dim sousTotal As Double
    Dim nomUtilisateur As String
    Dim prenomUtilisateur As String
    Dim emailUtilisateur As String
    Dim roleUtilisateur As String
    ' Variables suppl�mentaires qui seront utilis�es dans le code mais pas en interaction avec l'interface VBA
    Dim totalHT As Double
    Dim totalTTC As Double
    Dim clientId As Integer
    Dim factureId As Integer
    Dim moteurId As Integer
    
    ' Validation des champs de saisie avant l'initialisation des variables
    If Me.tNomClient.Value = "" Then
        MsgBox "Veuillez entrer le nom du client.", vbExclamation
        Exit Sub
    End If
    
    If Me.tAdresseClient.Value = "" Then
        MsgBox "Veuillez entrer l'adresse du client.", vbExclamation
        Exit Sub
    End If
    
    If Me.tContactClient.Value = "" Then
        MsgBox "Veuillez entrer le nom et pr�nom du contact.", vbExclamation
        Exit Sub
    End If
    
    If Me.tEmailClient.Value = "" Or InStr(1, Me.tEmailClient.Value, "@") = 0 Or InStr(1, Me.tEmailClient.Value, ".") = 0 Then
        MsgBox "Veuillez entrer une adresse email valide.", vbExclamation
        Exit Sub
    End If
    
    If Not IsDate(Me.tDateEmission.Value) Then
        MsgBox "Veuillez entrer une date valide.", vbExclamation
        Exit Sub
    End If
    
    If Me.tModalitePaiement.Value = "" Then
        MsgBox "Veuillez entrer une modalit� de paiement.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(Me.tDelaiPaiement.Value) Or Me.tDelaiPaiement.Value <= 0 Then
        MsgBox "Veuillez entrer un d�lai de paiement valide (nombre entier positif).", vbExclamation
        Exit Sub
    End If
    
    If Me.tDescription.Value = "" Then
        MsgBox "Veuillez entrer une description du produit.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(Me.tQuantite.Value) Or CInt(Me.tQuantite.Value) <= 0 Then
        MsgBox "Veuillez entrer une quantit� valide (nombre entier positif).", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(Me.tTauxDeTVA.Value) Or CDbl(Me.tTauxDeTVA.Value) < 0 Or CDbl(Me.tTauxDeTVA.Value) > 100 Then
        MsgBox "Veuillez entrer un taux de TVA valide (entre 0 et 100%).", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(Me.tPrixUnitaire.Value) Or CDbl(Me.tPrixUnitaire.Value) <= 0 Then
        MsgBox "Veuillez entrer un prix unitaire valide (nombre positif).", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(Me.tSousTotal.Value) Or CDbl(Me.tSousTotal.Value) <= 0 Then
        MsgBox "Le sous-total doit �tre un montant positif.", vbExclamation
        Exit Sub
    End If
    
    If Me.tNomUtilisateur.Value = "" Then
        MsgBox "Veuillez entrer le nom de l'utilisateur.", vbExclamation
        Exit Sub
    End If
    
    If Me.tPrenomUtilisateur.Value = "" Then
        MsgBox "Veuillez entrer le pr�nom de l'utilisateur.", vbExclamation
        Exit Sub
    End If
    
    If Me.tEmailUtilisateur.Value = "" Or InStr(1, Me.tEmailUtilisateur.Value, "@") = 0 Or InStr(1, Me.tEmailUtilisateur.Value, ".") = 0 Then
        MsgBox "Veuillez entrer une adresse email valide pour l'utilisateur.", vbExclamation
        Exit Sub
    End If
    
    If Me.tRoleUtilisateur.Value = "" Then
        MsgBox "Veuillez entrer le r�le de l'utilisateur.", vbExclamation
        Exit Sub
    End If

    ' Initialisation des variables du formulaire apr�s validation
    nomClient = Me.tNomClient.Value
    adresseClient = Me.tAdresseClient.Value
    contactClient = Me.tContactClient.Value
    emailClient = Me.tEmailClient.Value
    dateEmission = CDate(Me.tDateEmission.Value)
    modalitePaiement = Me.tModalitePaiement.Value
    delaisPaiement = CInt(Me.tDelaiPaiement.Value)
    descriptionProduit = Me.tDescription.Value
    quantiteProduit = CInt(Me.tQuantite.Value)
    tauxTVA = CDbl(Me.tTauxDeTVA.Value)
    prixUnitaire = CDbl(Me.tPrixUnitaire.Value)
    sousTotal = CDbl(Me.tSousTotal.Value)
    totalHT = sousTotal ' Total HT est �gal au sous-total pour ce cas
    totalTTC = sousTotal * (1 + tauxTVA / 100) ' Calcul du total TTC
    nomUtilisateur = Me.tNomUtilisateur.Value
    prenomUtilisateur = Me.tPrenomUtilisateur.Value
    emailUtilisateur = Me.tEmailUtilisateur.Value
    roleUtilisateur = Me.tRoleUtilisateur.Value
    
    ' Cr�ation et connexion � la base de donn�es
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=MysqlTestSource"
    
    On Error GoTo ErrorHandler
    
    ' On r�cup�re l'ID du client � partir de son nom
    clientId = GetClientId(conn, nomClient)
    
    If clientId = 0 Then
        MsgBox "Client non trouv� dans la base de donn�es.", vbExclamation
        Exit Sub
    End If
    
    ' Insertion des donn�es dans la table Factures et r�cup�ration de l'ID de la facture ins�r�e
  ' On r�cup�re l'ID du moteur associ� au client (le plus r�cent)
    moteurId = GetLatestMoteurId(conn, clientId)
    
    If moteurId = 0 Then
        MsgBox "Aucun moteur trouv� pour ce client.", vbExclamation
        Exit Sub
    End If
    
    Call InsererFacture(conn, clientId, dateEmission, totalTTC, "En attente", descriptionProduit, tauxTVA, totalHT, totalTTC, delaisPaiement)
    ' Insertion des donn�es dans la table Ventes
    factureId = RecupererFactureId(conn, clientId, dateEmission, totalHT, "En attente", descriptionProduit)
    
    ' Appel de la m�thode de g�n�ration de facture
    Call generationFacture(nomClient, contactClient, adresseClient, emailClient, factureId, clientId, dateEmission, descriptionProduit, quantiteProduit, prixUnitaire, totalHT, tauxTVA, totalTTC)

    'Call insertionVentes(conn, factureId, moteurId, quantiteProduit, prixUnitaire, dateEmission)
    
    ' Insertion des donn�es dans la table Utilisateurs
    'Call insertionUtilisateur(conn, nomUtilisateur, prenomUtilisateur, emailUtilisateur, roleUtilisateur, "Ecriture", dateEmission)
    
    
    ' Fermeture de la connexion et lib�ration des ressources
    conn.Close
    Set conn = Nothing
    
    MsgBox "Donn�es enregistr�es avec succ�s.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de l'enregistrement : " & Err.description, vbCritical
    conn.Close
    Set conn = Nothing
End Sub

' Fonction pour r�cup�rer l'ID du client � partir de son nom
Function GetClientId(conn As Object, nomClient As String) As Integer
    Dim rs As Object
    Dim strSQL As String
    strSQL = "SELECT client_id FROM clients WHERE nom = '" & nomClient & "'"
    Set rs = conn.Execute(strSQL)
    
    If Not rs.EOF Then
        GetClientId = rs.Fields("client_id").Value
    Else
        GetClientId = 0 ' Client non trouv�
    End If
    
    rs.Close
    Set rs = Nothing
End Function

' Fonction pour r�cup�rer l'ID du moteur le plus r�cent pour un client donn�
Function GetLatestMoteurId(conn As Object, clientId As Integer) As Integer
    Dim rs As Object
    Dim strSQL As String
    strSQL = "SELECT moteur_id FROM moteurs WHERE client_id = " & clientId & " ORDER BY date_achat DESC LIMIT 1"
    Set rs = conn.Execute(strSQL)
    
    If Not rs.EOF Then
        GetLatestMoteurId = rs.Fields("moteur_id").Value
    Else
        GetLatestMoteurId = 0 ' Aucun moteur trouv�
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Function RecupererFactureId(conn As Object, clientId As Integer, dateEmission As Date, montant As Double, etat As String, description As String) As Integer
    ' Construction de la requ�te SQL pour r�cup�rer l'ID de la facture bas�e sur les donn�es
    Dim strSQL As String
    strSQL = "SELECT facture_id FROM factures WHERE client_id = " & clientId & " AND date_emission = '" & Format(dateEmission, "yyyy-mm-dd") & _
             "' AND montant = " & Replace(CStr(montant), ",", ".") & " AND etat = '" & etat & "' AND description = '" & description & "'"
    
    ' Ex�cution de la requ�te SQL
    Dim rs As Object
    Set rs = conn.Execute(strSQL)
    
    ' V�rifie si un r�sultat a �t� renvoy�
    If Not rs.EOF Then
        RecupererFactureId = rs.Fields("facture_id").Value
    Else
        RecupererFactureId = 0 ' Retourne 0 si aucun ID trouv�
    End If
    
    rs.Close
    Set rs = Nothing
End Function
Sub InsererFacture(conn As Object, clientId As Integer, dateEmission As Date, montant As Double, etat As String, description As String, tauxTVA As Double, totalHT As Double, totalTTC As Double, delaisPaiement As Integer)
    ' Correction des s�parateurs d�cimaux pour SQL
    montant = Replace(CStr(montant), ",", ".")
    tauxTVA = Replace(CStr(tauxTVA), ",", ".")
    totalHT = Replace(CStr(totalHT), ",", ".")
    totalTTC = Replace(CStr(totalTTC), ",", ".")

    ' V�rification de la nullit� des valeurs importantes
    If IsNull(montant) Or IsNull(tauxTVA) Or IsNull(totalHT) Or IsNull(totalTTC) Then
        MsgBox "Erreur : Montant, Taux de TVA, Total HT ou Total TTC est vide ou null.", vbCritical
        Exit Sub
    End If
    
    ' Construction de la requ�te SQL
    Dim strSQL As String
    strSQL = "INSERT INTO factures (client_id, date_emission, montant, etat, description, taux_tva, total_ht, total_ttc, delais_paiement) VALUES (" & _
             clientId & ", '" & Format(dateEmission, "yyyy-mm-dd") & "', " & montant & ", '" & etat & "', '" & description & "', " & tauxTVA & ", " & totalHT & ", " & totalTTC & ", " & delaisPaiement & ")"
    
    ' Affichage de la requ�te SQL pour voir ce qui est envoy�
    MsgBox strSQL

    ' Ex�cution de la requ�te SQL
    conn.Execute (strSQL)
End Sub

' Proc�dure pour ins�rer une vente
Sub insertionVentes(conn As Object, factureId As Integer, moteurId As Integer, quantite As Integer, prixUnitaire As Double, dateVente As Date)
    Dim strSQL As String
    strSQL = "INSERT INTO ventes (facture_id, moteur_id, quantite, prix_unitaire, date_vente) VALUES (" & _
             factureId & ", " & moteurId & ", " & quantite & ", " & prixUnitaire & ", '" & Format(dateVente, "yyyy-mm-dd") & "')"
    MsgBox strSQL ' Affichage la requ�te SQL pour voir ce qui est envoye
    conn.Execute (strSQL)
End Sub

' Proc�dure pour ins�rer un utilisateur
Sub insertionUtilisateur(conn As Object, nom As String, prenom As String, email As String, role As String, permission As String, dateCreation As Date)
    Dim strSQL As String
    strSQL = "INSERT INTO utilisateurs (nom, prenom, email, role, permissions, date_creation) VALUES ('" & _
             nom & "', '" & prenom & "', '" & email & "', '" & role & "', '" & permission & "', '" & Format(dateCreation, "yyyy-mm-dd") & "')"
    MsgBox strSQL ' Affichage la requ�te SQL pour voir ce qui est envoye
    conn.Execute (strSQL)
End Sub

Private Sub bAnnuler_Click()
    ' Pour confirmer si l'utilisateur souhaite annuler l'op�ration
    Dim cancelConfirm As VbMsgBoxResult
    cancelConfirm = MsgBox("Voulez-vous vraiment annuler? Toutes les modifications non enregistr�es seront perdues.", vbYesNo + vbExclamation, "Annuler")
    
    If cancelConfirm = vbYes Then
        ' On r�initialise tous les champs du formulaire
        Me.tNomClient.Value = ""
        Me.tAdresseClient.Value = ""  ' V�rifiez bien que le nom du contr�le est correct
        Me.tContactClient.Value = ""
        Me.tEmailClient.Value = ""
        Me.tDateEmission.Value = ""
        Me.tModalitePaiement.Value = ""
        Me.tDelaiPaiement.Value = ""
        Me.tDescription.Value = ""
        Me.tQuantite.Value = ""
        Me.tTauxDeTVA.Value = ""
        Me.tPrixUnitaire.Value = ""
        Me.tSousTotal.Value = ""
        Me.tNomUtilisateur.Value = ""
        Me.tPrenomUtilisateur.Value = ""
        Me.tEmailUtilisateur.Value = ""
        Me.tRoleUtilisateur.Value = ""
        
        ' Retour � la premi�re page du formulaire
        MultiPage1.Value = 0
    End If
End Sub

Private Sub bPrecedentPage2_Click()
        MultiPage1.Value = 0 ' Page 1
End Sub

Private Sub bPrecedentPage3_Click()
    MultiPage1.Value = 1 ' Page 2
End Sub

Private Sub bPrecedentPage4_Click()
    MultiPage1.Value = 2 ' Page 3
End Sub
Private Sub bSuivantPage1_Click()
    If Me.tNomClient = "" Or Me.tAdresseClient = "" Or Me.tContactClient = "" Or Me.tDescription = "" Then
        MsgBox "Veuillez remplir tous les champs obligatoires sur cette page.", vbExclamation
    Else
        MultiPage1.Value = 1 ' Page 2
    End If
End Sub
Private Sub bSuivantPage2_Click()
    If Me.tDateEmission.Value = "" Or Me.tModalitePaiement.Value = "" Or Me.tDelaiPaiement.Value = "" Then
        MsgBox "Veuillez remplir tous les champs obligatoires sur cette page.", vbExclamation
    Else
        MultiPage1.Value = 2 ' Page 3
    End If
End Sub
Private Sub bSuivantPage3_Click()
    If Me.tDescription.Value = "" Or Me.tQuantite.Value = "" Or Me.tPrixUnitaire.Value = "" Or Me.tTauxDeTVA = "" Or Me.tSousTotal = "" Then
        MsgBox "Veuillez remplir tous les champs obligatoires sur cette page.", vbExclamation
    Else
        MultiPage1.Value = 3 ' Page 4
    End If
End Sub

'Validation en temps r�el des champs (�v�nements Exit)

'Valide que l'adresse est saisie.
Private Sub tAdresseClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.tAdresseClient.Value = "" Then
        MsgBox "Veuillez entrer une adresse valide.", vbExclamation
        Cancel = True
    End If
End Sub

' Valide que le contact est un nom et pr�nom valide (doit �tre une cha�ne de caract�res non vide).
Private Sub tContactClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(Me.tContactClient.Value) = "" Then
        MsgBox "Veuillez entrer un nom et pr�nom valide pour le contact.", vbExclamation
        Cancel = True
    End If
End Sub


'Valide que la date est correcte.
Private Sub tDateEmission_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(Me.tDateEmission.Value) Then
        MsgBox "Veuillez entrer une date valide.", vbExclamation
        Cancel = True
    End If
End Sub

'Valide que le d�lai de paiement est un entier positif.
Private Sub tDelaiPaiement_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Me.tDelaiPaiement.Value) Or Me.tDelaiPaiement.Value <= 0 Then
        MsgBox "Veuillez entrer un d�lai de paiement valide (nombre entier positif).", vbExclamation
        Cancel = True
    End If
End Sub

'Valide que l'email a une structure correcte.
Private Sub tEmailClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If InStr(1, Me.tEmailClient.Value, "@") = 0 Or InStr(1, Me.tEmailClient.Value, ".") = 0 Then
        MsgBox "Veuillez entrer une adresse email valide.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub tEmailUtilisateur_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If InStr(1, Me.tEmailUtilisateur.Value, "@") = 0 Or InStr(1, Me.tEmailUtilisateur.Value, ".") = 0 Then
        MsgBox "Veuillez entrer une adresse email valide.", vbExclamation
        Cancel = True
    End If
End Sub

'Valide que les modalit�s de paiement sont saisies.
Private Sub tModalitePaiement_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.tModalitePaiement.Value = "" Then
        MsgBox "Veuillez entrer une modalit� de paiement.", vbExclamation
        Cancel = True
    End If
End Sub

'Valide que le nom du client est saisi.
Private Sub tNomClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.tNomClient.Value = "" Then
        MsgBox "Veuillez entrer le nom du client.", vbExclamation
        Cancel = True
    End If
End Sub

'Valide que le nom de l'utilisateur est saisi.

Private Sub tNomUtilisateur_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.tNomUtilisateur.Value = "" Then
        MsgBox "Veuillez entrer le nom de l'utilisateur.", vbExclamation
        Cancel = True
    End If
End Sub

'Valide que le prix unitaire est un nombre positif.

Private Sub tPrixUnitaire_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Me.tPrixUnitaire.Value) Or Me.tPrixUnitaire.Value <= 0 Then
        MsgBox "Veuillez entrer un prix unitaire valide (nombre positif).", vbExclamation
        Cancel = True
    End If
End Sub

'Valide que la quantit� est un entier positif.

Private Sub tQuantite_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Me.tQuantite.Value) Or Me.tQuantite.Value <= 0 Then
        MsgBox "Veuillez entrer une quantit� valide (nombre entier positif).", vbExclamation
        Cancel = True
    End If
End Sub
'Valide que le r�le de l'utilisateur est saisi.
Private Sub tRoleUitilisateur_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.tRoleUtilisateur.Value = "" Then
        MsgBox "Veuillez entrer le r�le de l'utilisateur.", vbExclamation
        Cancel = True
    End If
End Sub

'Valide que le sous-total est un montant positif.

Private Sub tSousTotal_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Me.tSousTotal.Value) Or Me.tSousTotal.Value <= 0 Then
        MsgBox "Veuillez entrer un sous-total valide (nombre positif).", vbExclamation
        Cancel = True
    End If
End Sub

'Valide que le taux de TVA est un pourcentage valide.

Private Sub tTauxDeTVA_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Me.tTauxDeTVA.Value) Or Me.tTauxDeTVA.Value < 0 Or Me.tTauxDeTVA.Value > 100 Then
        MsgBox "Veuillez entrer un taux de TVA valide (0-100%).", vbExclamation
        Cancel = True
    End If
End Sub

Sub generationFacture(nomClient As String, contactClient As String, adresseClient As String, emailClient As String, factureId As Integer, clientId As Integer, dateEmission As Date, descriptionProduit As String, quantiteProduit As Integer, prixUnitaire As Double, totalHT As Double, tauxTVA As Double, totalTTC As Double)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Facture")
    
    ' Effacement des anciennes donn�es
    On Error Resume Next
    ws.Range("B11:B15").ClearContents ' Efface les informations du client
    ws.Range("F4:F6").ClearContents ' Efface les informations de la facture
    ' Efface les articles pr�c�dents
    Dim i As Integer
    For i = 1 To quantiteProduit
        ws.Range("B17").Offset(i, 0).ClearContents ' Num�ro de l'article
        ws.Range("B17").Offset(i, 1).ClearContents ' Description
        ws.Range("B17").Offset(i, 2).ClearContents ' Prix unitaire
        ws.Range("B17").Offset(i, 3).ClearContents ' Montant HT
        ws.Range("B17").Offset(i, 4).ClearContents ' Remise appliqu�e (aucune remise sp�cifique dans cet exemple)
    Next i
    ws.Range("E29:G33").ClearContents ' Efface les montants calcul�s (HT, TVA, TTC)
    On Error GoTo 0
    
    ' Ins�rer les nouvelles informations du client
    ws.Range("B11").Value = contactClient ' Contact du client
    ws.Range("B12").Value = nomClient ' Nom du client
    ws.Range("B13").Value = adresseClient ' Adresse du client
    ws.Range("B14").Value = emailClient ' Email du client

    ' Ins�rer les informations de la facture
    ws.Range("F4").Value = factureId ' ID de la facture
    ws.Range("F5").Value = clientId ' ID du client
    ws.Range("F6").Value = Format(dateEmission, "dd/mm/yyyy") ' Date de la facture

    ' Ins�rer les articles de la facture
    For i = 1 To quantiteProduit
        ws.Range("B17").Offset(i, 0).Value = i ' Num�ro de l'article
        ws.Range("B17").Offset(i, 1).Value = descriptionProduit ' Description
        ws.Range("B17").Offset(i, 2).Value = prixUnitaire ' Prix unitaire
        ws.Range("B17").Offset(i, 3).Value = prixUnitaire * i ' Montant HT
        ws.Range("B17").Offset(i, 4).Value = "-" ' Remise appliqu�e (aucune remise sp�cifique dans cet exemple)
    Next i

    ' Ins�rer les totaux
    ws.Range("E29").Value = totalHT ' Sous-total HT
    ws.Range("E31").Value = tauxTVA & "%" ' Taux de TVA
    ws.Range("E30").Value = totalTTC ' Total TTC

    ' Remise et total
    ws.Range("E32").Value = "0" ' Montant de la remise s'il y en a une
    ws.Range("E33").Value = totalTTC ' Solde d�

    ' Formatage des cellules
    ws.Range("E29:E33").NumberFormat = "#,##0.00 �" ' Format num�rique pour les montants
    
    MsgBox "Facture g�n�r�e avec succ�s.", vbInformation

End Sub
