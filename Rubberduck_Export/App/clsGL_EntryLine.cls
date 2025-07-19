VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGL_EntryLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Module de classe: clsGL_EntryLine - 2025-06-08 @ 07:28

Private pNoCompte As String
Private pDescription As String
Private pMontant As Currency

'Propriétés publiques pour accéder aux variables privées - Get & Let pour chaque variable

'NoCompte
Public Property Get NoCompte() As String
    NoCompte = pNoCompte
End Property

Public Property Let NoCompte(val As String)
    pNoCompte = val
End Property

'Description
Public Property Get description() As String
    description = pDescription
End Property

Public Property Let description(val As String)
    pDescription = val
End Property

'Montant (+ ou -)
Public Property Get montant() As Double
    montant = pMontant
End Property

Public Property Let montant(val As Double)
    pMontant = val
End Property

