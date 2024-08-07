VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16968
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fonction pour convertir toutes les entree de date en format souhaite
Private Function ConvertDateFormat(inputDate As String) As String
    Dim dt As Date
    dt = CDate(inputDate)
    ConvertDateFormat = Format(dt, "yyyy-mm-dd")
End Function

'Procedure pour Mettre à Jour Automatiquement les datas dans la Base de Données
Sub insertionIncident(moteurId As String, dateIncident As String, description As String, statut As String)
    Dim conn As Object
    Dim strSQL As String
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=MysqlTestSource"
    
    ' Conversion du format de la date
    dateIncident = ConvertDateFormat(dateIncident)
    
    ' Construction de la requête d'insertion
    strSQL = "INSERT INTO incidents (moteur_id, date_incident, description, statut) VALUES ('" & moteurId & "', '" & dateIncident & "', '" & description & "', '" & statut & "')"
    
    ' Afficher la requête SQL pour vérification
    MsgBox strSQL
    
    ' Exécution de la requête
    On Error GoTo SQLExecutionError
    conn.Execute strSQL
    
    ' Fermeture
    conn.Close
    Set conn = Nothing
    Exit Sub
    
SQLExecutionError:
    MsgBox "Erreur SQL : " & Err.description, vbCritical
    conn.Close
    Set conn = Nothing
End Sub


Private Sub bEnregistrer_Click()
    ' Déclaration des variables du formulaire
    Dim moteurId As String
    Dim description As String
    Dim dateIncident As String
    Dim statut As String
    
    ' Initialisation
    moteurId = Me.tMoteurId.Value
    description = Me.tDescription.Value
    dateIncident = Me.tDateIncident.Value
    statut = Me.tStatut.Value
    
    ' Vérification de la validation des données
    If moteurId = "" Or description = "" Or dateIncident = "" Or statut = "" Then
        MsgBox "Veuillez remplir tous les champs", vbExclamation
        Exit Sub ' Sortie si validation échoue
    End If
    
    ' Appel de la fonction pour insérer les données dans la base de données
    On Error GoTo ErrorHandler
    Call insertionIncident(moteurId, dateIncident, description, statut)
    MsgBox "Incident enregistré avec succès.", vbInformation
    
    ' Réinitialisation des champs du formulaire
    Me.tMoteurId.Value = ""
    Me.tDescription.Value = ""
    Me.tDateIncident.Value = ""
    Me.tStatut.Value = ""
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de l'enregistrement de l'incident : " & Err.description, vbCritical
End Sub

'Macro pour afficher le formulaire
Sub ShowIncidentForm()
    UserFormIncidents.Show
End Sub
'Procedure pour forcer la validation en temps reel, evitant toutes erreur de datas dans la bd en terme de structure

Private Sub tMoteurId_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Me.tMoteurId.Value) Then
        MsgBox "Veuillez entrer un ID de moteur valide.", vbExclamation
        Me.tMoteurId.SetFocus
    End If
End Sub


Private Sub tDateIncident_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(Me.tDateIncident.Value) Then
        MsgBox "Veuillez entrer une date valide, Referez vous a la base de donnees.", vbExclamation
        Me.tDateIncident.SetFocus
    End If
End Sub

Private Sub tStatus_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Me.tMoteurId.Value) Then
        MsgBox "Veuillez entrer un ID de moteur valide.", vbExclamation
        Me.tMoteurId.SetFocus
    End If
End Sub

Private Sub tDescription_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.tDescription.Value = "" Then
        MsgBox "Veuillez entrer une description valide.", vbExclamation
        Me.tDescription.SetFocus
    End If
End Sub


