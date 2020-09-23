VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_customize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customize"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_sounds 
      Caption         =   "Sounds"
      Height          =   1035
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Width           =   7275
      Begin VB.CheckBox chk_beep 
         Caption         =   "Play beeps through the PC Speaker as an audible alert"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   7035
      End
      Begin VB.CheckBox chk_speak 
         Caption         =   "Speak the new IP Address as an audible alert"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   7035
      End
   End
   Begin VB.CommandButton cmd_close 
      Caption         =   "Close"
      Height          =   315
      Left            =   6480
      TabIndex        =   1
      Top             =   2220
      Width           =   735
   End
   Begin VB.Frame fra_people 
      Caption         =   "People"
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      Begin VB.CommandButton cmd_change 
         Caption         =   "Change"
         Height          =   315
         Left            =   6420
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txt_bitmap 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   6255
      End
      Begin VB.CommandButton cmd_bitmap 
         Caption         =   "Bitmap"
         Height          =   315
         Left            =   6420
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin MSComDlg.CommonDialog com_dialog 
         Left            =   180
         Top             =   660
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Once you have selected a valid skin for Personal Frankwall, click Change to apply it."
         ForeColor       =   &H80000017&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frm_customize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************
'DISPLAY OPEN DIALOGUE
'************************************************************
Private Sub cmd_bitmap_Click()
    'CHANGE THE FILE FILTER TO BITMAPS SOLELY
    com_dialog.Filter = "Bitmap images (*.bmp)|*.bmp|"
    'SET THE START PATH THE THE APP PATH
    com_dialog.InitDir = App.Path
    'SHOW THE DIALOG
    com_dialog.ShowOpen
    'EVALUATE THE FILENAME AND RETRIEVE IF IT IS NOT EMPTY
    If com_dialog.FileName <> "" Then
        txt_bitmap = com_dialog.FileName
    End If
End Sub

'************************************************************
'APPLY THE SKIN MAP
'************************************************************
Private Sub cmd_change_Click()
    Dim keep As Integer     'MESSAGE BOX RETURN VARIABLE
    Dim previous As String  'THE PREVIOUS BITMAP
    'RETRIEVE THE PREVIOUS BITMAP
    previous = GetSetting(App.ProductName, "Customize", "Skin")
    Unload frm_main             'UNLOAD THE MAIN FORM
    reshaping = True            'MAKE THE FORM RESHAPE WHEN IT LOADS
    reshape_map = txt_bitmap    'SET THE NEW BITMAP
    Load frm_main               'LOAD AND SHAPE THE FORM
    'PROMPT USER TO KEEP THE BITMAP
    Me.Hide
    keep = MsgBox("Do you wish to keep the current skin?", vbYesNo, "Skin change complete")
    Me.Show
    If keep = vbYes Then
        'SAVE THE SETTING
        SaveSetting App.ProductName, "Customize", "Skin", txt_bitmap
    Else
        Unload frm_main             'UNLOAD THE MAIN FORM
        reshaping = True            'MAKE THE FORM RESHAPE WHEN IT LOADS
        reshape_map = previous      'SET THE BITMAP BACK TO ITS PREVOUS ONE
        Load frm_main               'LOAD AND RESHAPE THE FORM
    End If
End Sub

'************************************************************
'CLOSE THIS FORM
'************************************************************
Private Sub cmd_close_Click()
    SaveSetting App.ProductName, "Customize", "Sounds - beep", chk_beep
    SaveSetting App.ProductName, "Customize", "Sounds - speak", chk_speak
    play_speak = chk_speak
    play_beep = chk_beep
    Unload Me   'UNLOAD THE CUSTOMIZE FORM
End Sub

'************************************************************
'LOAD THIS FORM
'************************************************************
Private Sub Form_Load()
    chk_beep = play_beep
    chk_speak = play_speak
    Me.Show
End Sub
