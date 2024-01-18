VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormAjout 
   Caption         =   "Ajouter un élève"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   OleObjectBlob   =   "UserFormAjout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormAjout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButtonAjouter_Click()
    ' Appeler la procédure AjouterPrenomNom avec les valeurs des TextBox
    AjouterPrenomNom TextBoxPrenom.Value, TextBoxNom.Value
    
End Sub

Private Sub CommandButtonAnnuler_Click()
    ' Cacher le UserForm sans ajouter les données
    Unload Me
End Sub
