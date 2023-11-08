Attribute VB_Name = "Admin_Macros"
Option Explicit

Sub Admin_SetPlatformColor()
    ActiveCell.Interior.Color = Admin.Shapes(Application.Caller).Fill.ForeColor.RGB
    Admin.Shapes("ColorPalette").Visible = msoFalse
End Sub

