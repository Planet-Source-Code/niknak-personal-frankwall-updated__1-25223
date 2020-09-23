VERSION 5.00
Begin VB.Form frm_menus 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MENUS"
   ClientHeight    =   135
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   1515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   135
   ScaleWidth      =   1515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu men_menu 
      Caption         =   "Menu"
      Begin VB.Menu men_customize 
         Caption         =   "Customize"
      End
      Begin VB.Menu men_spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu men_about 
         Caption         =   "About"
      End
      Begin VB.Menu men_spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu men_exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frm_menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'************************************************************
'MENU FUNCTIONS ARE IN A SEPARATE FORM DUE TO THEM BECOMING
'VISIBLE ON THE MAIN FORM, EVEN THOUGH IT HAS NO TITLE BAR
'ALSO KEEPS THEM SEPARATE AND TIDIES CODE UP A TAD
'************************************************************
'SPLASH SCREEN
Private Sub men_about_Click()
    'SHOW SPLASH SCREEN
    frm_splash.Show
End Sub

'CUSTOMIZE MENU
Private Sub men_customize_Click()
    'LOAD THE CUSTOMIZE FORM
    Load frm_customize
End Sub

'MAIN UNLOAD INTERFACE, DONT UNLOAD ANYWHERE ELSE!!!
Private Sub men_exit_Click()
    'UNLOAD ALL OTHER FORMS
    Unload frm_splash
    Unload frm_winsock
    'UNLOAD FRANK
    Unload frm_main
    'UNLOAD ME
    End
End Sub
'************************************************************

