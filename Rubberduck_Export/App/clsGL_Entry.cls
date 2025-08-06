VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGL_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Module de classe: clsGL_Entry - 2025-06-08 @ 07:35

Public pNoEcriture As Long
Public pDateEcriture As Date
Public pDescription As String
Public pSource As String

Public pLignes As Collection

'Constructeur (s'exécute à la création)
Private Sub Class_Initialize()
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
Public Property Get description() As String
    description = pDescription
End Property

Public Property Let description(val As String)
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
Public Property Get source() As String
    source = pSource
End Property

Public Property Let source(val As String)
    pSource = val
End Property

'Accès aux lignes (Collection d'objets clsGL_EntryLine)
Public Property Get lignes() As Collection
    Set lignes = pLignes
End Property

'Méthode pour ajouter une ligne à l'écriture
Public Sub AjouterLigne(noCompte As String, description As String, montant As Currency, autreRemarque As String)

    Dim l As clsGL_EntryLine
    Set l = New clsGL_EntryLine
    l.noCompte = noCompte
    l.description = description
    l.montant = montant
    l.autreRemarque = autreRemarque
    pLignes.Add l
    
End Sub
