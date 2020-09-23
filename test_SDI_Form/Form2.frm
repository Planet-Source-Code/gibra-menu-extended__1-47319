VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2 - Adesso fai clic destro sul form..."
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    ' se non è attiva la classe MenuEx
    If objMenuEx Is Nothing Then
      If Shift And vbCtrlMask Then
        ' questo menu è nel menu Modifica
        SDIForm.PopupMenu SDIForm.mnuBar(2)
      Else
        ' questo menu è sulla barra
        SDIForm.PopupMenu SDIForm.mnuHidden
      End If
    Else
      If Shift And vbCtrlMask Then
        ' questo menu è nel menu Modifica
        Call objMenuEx.PopupMenu(SDIForm, SDIForm.mnuBar(2))
      Else
        ' questo menu è sulla barra
        Call objMenuEx.PopupMenu(SDIForm, SDIForm.mnuHidden)
      End If
    End If
  End If
  
End Sub


