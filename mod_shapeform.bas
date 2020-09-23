Attribute VB_Name = "mod_shapeform"
Option Explicit

'REGION HANDLING
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Const RGN_OR = 2
Public Const RGN_XOR = 3

'IMAGE HANDLING
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'MOVEMENT HANDLING
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public reshaping As Boolean
Public reshape_map As String

'GLOBAL VARIABLES
Private FORM_WIDTH As Long      'FORM WIDTH
Private FORM_HEIGHT As Long     'FORM HEIGHT

'PRIVATE VARIABLES
Private DC As Long              'FORM DC HANDLE
Private BMP As Long             'BITMAP HANDLE
Private Pix As Long             'CURRENT PIXEL COLOUR
Private rgnInv As Long          'REGION JUNK
Private rgn As Long             'REGION JUNK
Private rgnTotal As Long        'REGION JUNK
    
Public Function CreateRegionFromFile(shape_form As Form, shape_map As Image, strFile, BGColor As Long) As Long
    Dim height_counter As Integer 'WIDTH COUNTER
    Dim width_counter As Integer 'HEIGHT COUNTER
    
    'LOAD THE PICTURE INTO AN IMAGE BOX ONTO A NON BORDERED FORM!
    shape_map.Picture = LoadPicture(strFile)
    'WRAP THE FORM TO THE IMAGE BOX
    shape_form.Width = shape_map.Width
    shape_form.Height = shape_map.Height
    'MAKE THE IMAGE INVISIBLE
    shape_map.Visible = False
    
    'LOAD THE IMAGE INTO THE BACK OF THE FORM
    shape_form.Picture = LoadPicture(strFile)
    
    'SET THE SCALE MODE TO PIXELS
    shape_form.ScaleMode = vbPixels
    
    'SET THE TWO VARIABLES
    'FW - FORM WIDTH
    'FH - FORM HEIGHT
    'TO THE WIDTH AND HEIGHT OF THE SCALED FORM
    FORM_WIDTH = shape_form.ScaleWidth
    FORM_HEIGHT = shape_form.ScaleHeight
    
    'CREATE COMPATIBLE DISPLAY CONTEXT
    DC = CreateCompatibleDC(shape_form.hdc)
    'LOAD THE FORM BITMAP INTO BMP
    BMP = SelectObject(DC, LoadPicture(strFile))
    
    'REGION SETUP
    rgnTotal = CreateRectRgn(0, 0, FORM_WIDTH, FORM_HEIGHT)
    rgnInv = CreateRectRgn(0, 0, FORM_WIDTH, FORM_HEIGHT)
    CombineRgn rgnTotal, rgnTotal, rgnTotal, RGN_XOR
    
    'GO THROUGH THE FORM AND REMOVE ALL BACKGROUND COLOURED PIXELS
    For height_counter = 0 To FORM_HEIGHT
        For width_counter = 0 To FORM_WIDTH
            Pix = GetPixel(DC, width_counter, height_counter)
            If Pix = BGColor Then
                rgn = CreateRectRgn(width_counter, height_counter, width_counter + 1, height_counter + 1)
                CombineRgn rgnTotal, rgnTotal, rgn, RGN_OR
                DeleteObject rgn
            End If
        Next width_counter
    Next height_counter

    CombineRgn rgnTotal, rgnTotal, rgnInv, RGN_XOR
    CreateRegionFromFile = rgnTotal
End Function
