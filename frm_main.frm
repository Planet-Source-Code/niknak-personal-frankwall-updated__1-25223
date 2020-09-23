VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form frm_main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   510
   ClientLeft      =   -90
   ClientTop       =   -660
   ClientWidth     =   4245
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin HTTSLibCtl.TextToSpeech voc_speak 
      Height          =   375
      Left            =   2160
      OleObjectBlob   =   "frm_main.frx":0CCA
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   2760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tim_flash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3360
      Top             =   60
   End
   Begin VB.Timer tim_check 
      Interval        =   1
      Left            =   3780
      Top             =   60
   End
   Begin VB.Image img_frank 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frm_main.frx":0CEE
      MousePointer    =   99  'Custom
      Top             =   60
      Width           =   315
   End
   Begin VB.Label lbl_alert 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image img_shapemap 
      Height          =   510
      Left            =   0
      Picture         =   "frm_main.frx":19F0
      Top             =   0
      Width           =   4245
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************
'API DECLARATIONS
'************************************************************
'USED TO KEEP FORM ONTOP
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
'PC SPEAKER BEEP
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'************************************************************
'PRIVATE VARIABLES
'************************************************************
Dim firstload As Boolean        'PF'S FIRST LOAD
Dim current_ipadd As String     'YOUR CURRENT IP ADDRESS
Dim hidden As Boolean           'HIDDEN FLAG
'************************************************************

'************************************************************
'THE MAIN LOAD, ALSO RESHAPES FRANK
'************************************************************
Private Sub Form_Load()
    '********************************
    'LOAD SOUND SETTINGS
    play_speak = Val(GetSetting(App.ProductName, "Customize", "Sounds - speak"))
    play_beep = Val(GetSetting(App.ProductName, "Customize", "Sounds - beep"))
    '********************************
    Dim face As Variant
    If Not firstload Then
        check_ipadd
        If GetSetting(App.ProductName, "Customize", "Skin") <> "" Then
            reshaping = True
            reshape_map = GetSetting(App.ProductName, "Customize", "Skin")
        End If
        firstload = True
    End If
    If reshaping Then
        load_window App.ProductName
        img_shapemap = LoadPicture(reshape_map)
        SavePicture img_shapemap.Picture, App.Path & "\shapemap.tmp"
        face = CreateRegionFromFile(Me, img_shapemap, App.Path & "\shapemap.tmp", RGB(0, 255, 0))
        SetWindowRgn Me.hwnd, face, True
        hideme
        Me.Visible = True
        reshaping = False
    Else
        load_window App.ProductName
        SavePicture img_shapemap.Picture, App.Path & "\shapemap.tmp"
        face = CreateRegionFromFile(Me, img_shapemap, App.Path & "\shapemap.tmp", RGB(0, 255, 0))
        SetWindowRgn Me.hwnd, face, True
        hideme
        Me.Visible = True
    End If
    current_ipadd = ""
    check_ipadd
End Sub

'************************************************************
'MAIN UNLOAD
'************************************************************
Private Sub Form_Unload(Cancel As Integer)
    'STOP THE TIMERS
    tim_flash.Enabled = False
    tim_check.Enabled = False
    'SAVE FRANKS YPOS
    save_window App.ProductName, Me.Top
End Sub

'************************************************************
'TOGGLE THE HIDDEN FLAG AND UPDATE THE BAR
'************************************************************
Private Sub img_frank_DblClick()
    'STOP FLASHING TIMER
    tim_flash.Enabled = False
    'MAKE SURE LABEL IS IVISBLE
    lbl_alert.Visible = True
    'TOGGLE FRANKS HIDDEN FLAG
    If hidden Then
        showme
    Else
        hideme
    End If
End Sub

'************************************************************
'THIS SUB MOVES THE BAR AND DISPLAYS THE POPUP
'************************************************************
Private Sub img_frank_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'STOP SPEAKING
    voc_speak.StopSpeaking
    'STOP FLASHING TIMER
    tim_flash.Enabled = False
    'MAKE SURE LABEL IS IVISBLE
    lbl_alert.Visible = True
    'EVALUATE BUTTON
    If Button = vbLeftButton Then
        'MOVE FRANK
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    Else
        'DISPLAY POPUP
        PopupMenu frm_menus.men_menu
    End If
    'MOVE FRANK BACK TO OLD X POSITION
    If hidden Then
        hideme
    Else
        showme
    End If
End Sub

'************************************************************
'THIS SUB SAVES THE WINDOW'S Y POSITION
'************************************************************
Public Sub save_window(window As String, save_top As Long)
    'SAVE FRANKS POSITION
    SaveSetting App.ProductName, "windows", window, "SAVED"
    SaveSetting App.ProductName, "windows", window & " top", save_top
End Sub

'************************************************************
'THIS SUB LOADS THE WINDOWS Y POSITION
'************************************************************
Public Sub load_window(window As String)
    Dim win_top As Long 'USED FOR FRANKS TOP POSITION
    If GetSetting(App.ProductName, "windows", window) = "SAVED" Then
        'OBTAIN SAVED VALUE
        win_top = Val(GetSetting(App.ProductName, "windows", window & " top"))
    Else
        'USE 0 AS DEFAULT
        win_top = 0
    End If
    'CHECK THE TOP VALUE AGAINST THE SCREEN HEIGHT, AND IF NEED, MOVE TO DEFAULT
    If win_top > Screen.Height Then win_top = 0
    'MOVE FRANK TO THE OBTAINED POSITION OR DEFAULT
    Me.Top = win_top
End Sub

'************************************************************
'THIS TIMER BRINGS THE BAR TO THE TOP AND CHECKS THE IP
'************************************************************
Private Sub tim_check_Timer()
    'PUT FRANK ONTOP
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    'CHECK THE IP ADDRESS
    check_ipadd
End Sub

'************************************************************
'THIS SUB LOADS UP THE WINSOCK FORM TO OBTAIN FRESH LOCAL
'DETAILS AND RETRIEVES THE IP ADDRESS, THEN COMPARES IT
'AGAINST THE SAVED ONE AND RESPONDS AS NECESSARY
'************************************************************
Private Sub check_ipadd()
    'LOAD UP A NEW WINSOCK OBJECT TO OBTAIN THE NEW ADDRESS
    Load frm_winsock
    'CHECK THE ADDRESS AGAINST THE SAVE ADDRESS
    If frm_winsock.win_frank.LocalIP <> current_ipadd Then
        'START THE FLASH TIMER
        tim_flash.Enabled = True
        'SAVE THE NEW ADDRESS
        current_ipadd = frm_winsock.win_frank.LocalIP
        'SHOW FRANK WITH THE NEW ADDRESS
        showme "IP address is " & current_ipadd
        'SPEAK NEW ADDRESS IF WANTED
        If play_speak Then voc_speak.Speak current_ipadd
    End If
    'UNLOAD THE WINSOCK FORM
    Unload frm_winsock
End Sub

'************************************************************
'THIS TIMER FLASHES THE LABEL AND PLAYS THE BEEP SOUND
'BUT ONLY IF REQUIRED
'************************************************************
Private Sub tim_flash_Timer()
    'TOGGLE THE LABELS VISIBILITY FLAG
    lbl_alert.Visible = Not lbl_alert.Visible
    'PLAY BEEP ALERT IF WANTED
    If play_beep Then
        If lbl_alert.Visible Then
            'HIGH TONE
            Beep 200, 100
        Else
            'LOW TONE
            Beep 400, 100
        End If
    End If
End Sub

'************************************************************
'THE FOLLOWING TWO SUB ROUTINES SHOW OR HIDE FRANK
'THE SHOW FUNCTION HAS AN OPTION MESSAGE PARAMETER WHICH
'ENABLES YOU TO DISPLAY ANYTHING ON THE FRANK BAR
'************************************************************
'HIDE FRANKBAR
Private Sub hideme()
    Dim next_pos As Long    'USED FOR MOVEMENT LOOP
    'CHANGE TO TOOLTIP FOR THE FACE PIC
    img_frank.ToolTipText = "Double click to open"
    'MOVE THE BAR INTO POSITION
    For next_pos = Me.Left To (Screen.Width - 187) Step 1
        Me.Left = next_pos
    Next next_pos
    'FORCE BAR TO FINAL POSITION
    Me.Left = Screen.Width - 187
    'CHANGE THE HIDDEN FLAG
    hidden = True
End Sub

'SHOW FRANKBAR
Private Sub showme(Optional message As String)
    Dim next_pos As Long    'USED FOR MOVEMENT LOOP
    'USE MESSAGE IF PASSED
    If message <> "" Then
        lbl_alert.Caption = message
    End If
    'CHANGE TO TOOLTIP FOR THE FACE PIC
    img_frank.ToolTipText = "Double click to close"
    'MOVE THE BAR INTO POSITION
    For next_pos = Me.Left To (Screen.Width - Me.Width) Step -4
        Me.Left = next_pos
    Next next_pos
    'FORCE BAR TO FINAL POSITION
    Me.Left = Screen.Width - Me.Width
    'CHANGE THE HIDDEN FLAG
    hidden = False
End Sub
'************************************************************

