VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGL_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Module de classe: cGL_Entry - 2025-06-08 @ 07:35

Private pNoEcriture As Long
Private pDateEcriture As Date
Private pDescription As String
Private pSource As String
Private pAutreRemarque As String

Private pLignes As Collection

'Constructeur (s'exécute à la création)
Sub Class_Initialize()
    Set pLignes = New Collection
End Sub

'Propriétés communes

'NoEcriture
Public Property Get NoEcriture() As Long
    NoEcriture = pNoEcriture
End Property

Public Property Let NoEcriture(val As Long)
    pNoEcriture = val
End Property

'Description
Public Property Get Description() As String
    Description = pDescription
End Property

Public Property Let Description(val As String)
    pDescription = val
End Property

'Date écriture
Public Property Get DateEcriture() As Date
    DateEcriture = pDateEcriture
End Property

Public Property Let DateEcriture(val As Date)
    pDateEcriture = val
End Property

'Source
Public Property Get Source() As String
    Source = pSource
End Property

Public Property Let Source(val As String)
    pSource = val
End Property

'Autre remarque
Public Property Get AutreRemarque() As String
    AutreRemarque = pAutreRemarque
End Property

Public Property Let AutreRemarque(val As String)
    pAutreRemarque = val
End Property

'Accès aux lignes (Collection d'objets cGL_EntryLine)
Public Property Get Lignes() As Collection
    Set Lignes = pLignes
End Property

'Méthode pour ajouter une ligne à l'écriture
Public Sub AjouterLigne(NoCompte As String, Description As String, Montant As Currency)

    Dim l As cGL_EntryLine
    Set l = New cGL_EntryLine
    l.NoCompte = NoCompte
    l.Description = Description
    l.Montant = Montant
    pLignes.Add l
    
End Sub
