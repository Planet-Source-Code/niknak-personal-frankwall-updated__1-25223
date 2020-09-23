VERSION 5.00
Begin VB.Form frm_splash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Personal Frankwall"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image img_splash 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Left            =   120
      Picture         =   "frm_splash.frx":0000
      Top             =   120
      Width           =   3810
   End
End
Attribute VB_Name = "frm_splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'UNLOAD THE SPLASH FORM
'************************************************************
Private Sub img_splash_Click()
    Unload Me
End Sub
'************************************************************

