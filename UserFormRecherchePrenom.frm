VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormRecherchePrenom 
   Caption         =   "Rechercher Prénom"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4545
   OleObjectBlob   =   "UserFormRecherchePrenom.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormRecherchePrenom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Resultat As String

Private Sub CommandButtonRechercher_Click()
    ' Mettre à jour la variable Resultat avec la valeur du TextBox
    Resultat = TextBoxNumeroEtudiant.Value
    
    ' Exécuter la recherche avec le numéro d'étudiant
    RechercherEtAfficherPrenom Resultat
End Sub

Private Sub CommandButtonAnnuler_Click()
    ' Cacher le UserForm
    Unload Me
End Sub
