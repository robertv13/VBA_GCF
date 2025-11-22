Attribute VB_Name = "modTests"
Option Explicit

Sub zz_RunAllTestsUnitaires()

    Debug.Print String(36, "-")
    Debug.Print "Début de la suite de tests unitaires"
    Debug.Print String(36, "-")
    
    'Tests du module modAuditVBA
    Call TU_Fn_ExtractProcName
    Call TU_FnAllerVersCode
    
    'Tests du module modFunctions
    Call TU_FnCompleteLaDate
    
    Debug.Print "Fin de la suite de tests unitaires"
    Debug.Print String(34, "-")
    
    Call Fn_AllerVersCode("modTests", "RunAllTestsUnitaires")

End Sub

