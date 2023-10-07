VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplashScreen 
   Caption         =   "SplashUserForm"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSplashScreen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    HideTitleBar Me
End Sub

Private Sub UserForm_Activate()
    Application.Wait (Now + TimeValue("00:00:02"))
    frmSplashScreen.Label3.Caption = "Loading Data..."
    frmSplashScreen.Repaint
    Application.Wait (Now + TimeValue("00:00:02"))
    frmSplashScreen.Label3.Caption = "Creating Forms..."
    frmSplashScreen.Repaint
    Application.Wait (Now + TimeValue("00:00:02"))
    frmSplashScreen.Label3.Caption = "Opening..."
    frmSplashScreen.Repaint
    Application.Wait (Now + TimeValue("00:00:02"))
    Unload frmSplashScreen
End Sub

