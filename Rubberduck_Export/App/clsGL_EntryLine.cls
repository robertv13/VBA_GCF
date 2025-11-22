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
Private pAutreRemarque As String


'Propriétés publiques pour accéder aux variables privées - Get & Let pour chaque variable

Public Property Get noCompte() As String

    noCompte = pNoCompte
    
End Property

Public Property Let noCompte(val As String)

    pNoCompte = val
    
End Property

Public Property Get description() As String

    description = pDescription
    
End Property

Public Property Let description(val As String)

    pDescription = val
    
End Property

Public Property Get montant() As Double

    montant = pMontant
    
End Property

Public Property Let montant(val As Double)

    pMontant = val
    
End Property

'Autre remarque
Public Property Get autreRemarque() As String

    autreRemarque = pAutreRemarque
    
End Property

Public Property Let autreRemarque(val As String)

    pAutreRemarque = val
    
End Property


