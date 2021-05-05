VERSION 4.00
Begin VB.Form frmRGB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peter Chapman's Colour Slider 2.9"
   ClientHeight    =   3630
   ClientLeft      =   2580
   ClientTop       =   4035
   ClientWidth     =   5775
   Height          =   4140
   Icon            =   "RGB.frx":0000
   Left            =   2520
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5775
   Top             =   3585
   Width           =   5895
   Begin VB.CommandButton cmdSwap 
      Caption         =   "&Swap Foreground and Background"
      Height          =   375
      Left            =   720
      TabIndex        =   20
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtOLEColour 
      Height          =   285
      Left            =   4680
      MaxLength       =   8
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox chkLowerCase 
      Caption         =   "Generate &Lower Case Colour Code"
      Height          =   195
      Left            =   2640
      TabIndex        =   24
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "&Keep Window On Top"
      Height          =   195
      Left            =   360
      TabIndex        =   23
      Top             =   3360
      Width           =   2055
   End
   Begin VB.OptionButton optForeground 
      Caption         =   "&Foreground"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   1785
      Width           =   1215
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "B&ackground"
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   1785
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtHTMLColour 
      Height          =   285
      Left            =   4680
      MaxLength       =   7
      TabIndex        =   15
      Top             =   1770
      Width           =   975
   End
   Begin VB.TextBox txtGreen 
      Height          =   285
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtRedHex 
      Height          =   285
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   11
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtGreenHex 
      Height          =   285
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtBlueHex 
      Height          =   285
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtBlue 
      Height          =   285
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtRed 
      Height          =   285
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   2715
      Width           =   1335
   End
   Begin VB.Label lblGeneral 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&OLE Colour"
      Height          =   195
      Index           =   6
      Left            =   3615
      TabIndex        =   16
      Top             =   2190
      Width           =   810
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      Caption         =   "The quick brown fox jumps over the very lazy dog"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   195
      TabIndex        =   21
      Top             =   2640
      Width           =   3945
   End
   Begin ComctlLib.Slider SliderB 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1275
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   661
      _StockProps     =   64
      Max             =   255
   End
   Begin ComctlLib.Slider SliderG 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   795
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   661
      _StockProps     =   64
      Max             =   255
   End
   Begin ComctlLib.Slider SliderR 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   330
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   661
      _StockProps     =   64
      Max             =   255
   End
   Begin VB.Label lblGeneral 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "HTML &Colour"
      Height          =   195
      Index           =   5
      Left            =   3510
      TabIndex        =   14
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label lblGeneral 
      Alignment       =   2  'Center
      Caption         =   "&Hex"
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblGeneral 
      Alignment       =   2  'Center
      Caption         =   "&Dec"
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblGeneral 
      Caption         =   "&Blue"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblGeneral 
      Caption         =   "&Green"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblGeneral 
      Caption         =   "&Red"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmRGB"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' Require variable declaration
Option Explicit

' The other colour (background/foreground)
Dim OtherColour As String

' True if the colour is being set via HTML text
Dim SetHTMLColourExecuting As Boolean

' True if the colour is being set via OE number
Dim SetOLEColourExecuting As Boolean

' Change to lower case checkbox click event handler
Private Sub chkLowerCase_Click()
    If chkLowerCase.Value Then
        txtHTMLColour.Text = LCase(txtHTMLColour.Text)
    Else
        txtHTMLColour.Text = UCase(txtHTMLColour.Text)
    End If
End Sub

' Keep window on top checkbox click event handler
Private Sub chkOnTop_Click()
    If chkOnTop.Value Then
        SetTopMostWindow Me.hwnd, True
    Else
        SetTopMostWindow Me.hwnd, False
    End If
End Sub

' Exit button click event handler
Private Sub cmdExit_Click()
    Unload Me
End Sub

' Swap Background and Foreground Colours button click event handler
Private Sub cmdSwap_Click()
    If optForeground.Value Then
        lblMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
        optForeground_Click
        lblMain.ForeColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    Else
        lblMain.ForeColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
        optForeground_Click
        lblMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    End If
End Sub

' Form load event handler
Private Sub Form_Load()
    Me.Top = GetSetting("Peter Chapman", "Colour Slider", "Top", (Screen.Height - Me.Height) / 2)
    Me.Left = GetSetting("Peter Chapman", "Colour Slider", "Left", (Screen.Width - Me.Width) / 2)
    chkOnTop.Value = GetSetting("Peter Chapman", "Colour Slider", "OnTop", 0)
    chkLowerCase.Value = GetSetting("Peter Chapman", "Colour Slider", "LowerCase", 0)
    optForeground.Value = GetSetting("Peter Chapman", "Colour Slider", "Foreground", False)
    optBackground.Value = GetSetting("Peter Chapman", "Colour Slider", "Background", True)
    txtRed.Text = GetSetting("Peter Chapman", "Colour Slider", "Red", 255)
    txtGreen.Text = GetSetting("Peter Chapman", "Colour Slider", "Green", 255)
    txtBlue.Text = GetSetting("Peter Chapman", "Colour Slider", "Blue", 255)
    OtherColour = GetSetting("Peter Chapman", "Colour Slider", "OtherColour", "#000000")
    If optForeground.Value Then
        optBackground.Value = True
        optForeground.Value = True
    Else
        optForeground.Value = True
        optBackground.Value = True
    End If
End Sub

' Form unload event handler
Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Peter Chapman", "Colour Slider", "Red", txtRed.Text
    SaveSetting "Peter Chapman", "Colour Slider", "Green", txtGreen.Text
    SaveSetting "Peter Chapman", "Colour Slider", "Blue", txtBlue.Text
    SaveSetting "Peter Chapman", "Colour Slider", "Foreground", optForeground.Value
    SaveSetting "Peter Chapman", "Colour Slider", "Background", optBackground.Value
    SaveSetting "Peter Chapman", "Colour Slider", "OtherColour", OtherColour
    SaveSetting "Peter Chapman", "Colour Slider", "OnTop", chkOnTop.Value
    SaveSetting "Peter Chapman", "Colour Slider", "LowerCase", chkLowerCase.Value
    SaveSetting "Peter Chapman", "Colour Slider", "Top", Me.Top
    SaveSetting "Peter Chapman", "Colour Slider", "Left", Me.Left
    End
End Sub

' The background option button click event handler
Private Sub optBackground_Click()
    optForeground_Click
End Sub

' The foreground option button click event handler
Private Sub optForeground_Click()
    Dim sCurrentColour As String
    sCurrentColour = txtHTMLColour.Text
    txtHTMLColour.Text = OtherColour
    OtherColour = sCurrentColour
End Sub

' Blue scrollbar change event handler
Private Sub SliderB_Change()
    SliderB_Scroll
End Sub

' Blue scrollbar scroll event handler
Private Sub SliderB_Scroll()
    txtBlue.Text = SliderB.Value
    txtBlueHex.Text = Hex(SliderB.Value)
    If optBackground.Value Then
        lblMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    Else
        lblMain.ForeColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    End If
End Sub

' Green scrollbar change event handler
Private Sub SliderG_Change()
    SliderG_Scroll
End Sub

' Green scrollbar scroll event handler
Private Sub SliderG_Scroll()
    txtGreen.Text = SliderG.Value
    txtGreenHex.Text = Hex(SliderG.Value)
    If optBackground.Value Then
        lblMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    Else
        lblMain.ForeColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    End If
End Sub

' Red scrollbar change event handler
Private Sub SliderR_Change()
    SliderR_Scroll
End Sub

' Red scrollbar scroll event handler
Private Sub SliderR_Scroll()
    txtRed.Text = SliderR.Value
    txtRedHex.Text = Hex(SliderR.Value)
    If optBackground.Value Then
        lblMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    Else
        lblMain.ForeColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    End If
End Sub

' Decimal blue textbox change event handler
Private Sub txtBlue_Change()
    On Error GoTo txtBlueErr
    SliderB.Value = txtBlue.Text
    If txtBlueHex.Text <> Hex(txtBlue.Text) Then txtBlueHex.Text = Hex(txtBlue.Text)
    If optBackground.Value Then
        lblMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    Else
        lblMain.ForeColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    End If
txtBlueErr:
End Sub

' Hexadecimal blue textbox change event handler
Private Sub txtBlueHex_Change()
    txtBlue.Text = Dec(txtBlueHex.Text)
    If txtBlueHex.Text <> Mid(txtHTMLColour.Text, 6, 2) Then SetHTMLAndOLEColour
End Sub

' Decimal green textbox change event handler
Private Sub txtGreen_Change()
    On Error GoTo txtGreenErr
    SliderG.Value = txtGreen.Text
    If txtGreenHex.Text <> Hex(txtGreen.Text) Then txtGreenHex.Text = Hex(txtGreen.Text)
    If optBackground.Value Then
        lblMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    Else
        lblMain.ForeColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    End If
txtGreenErr:
End Sub

' Hexadecimal green textbox change event handler
Private Sub txtGreenHex_Change()
    txtGreen.Text = Dec(txtGreenHex.Text)
    If txtGreenHex.Text <> Mid(txtHTMLColour.Text, 4, 2) Then SetHTMLAndOLEColour
End Sub

' HTML colour textbox change event handler
Private Sub txtHTMLColour_Change()
    Dim hexColour As String
    If Left(txtHTMLColour.Text, 1) = "#" And Not SetHTMLColourExecuting Then
        hexColour = txtHTMLColour.Text
        If Len(hexColour) < 7 Then hexColour = hexColour + String(7 - Len(hexColour), "0")
        txtRedHex.Text = UCase(Mid(hexColour, 2, 2))
        txtGreenHex.Text = UCase(Mid(hexColour, 4, 2))
        txtBlueHex.Text = UCase(Mid(hexColour, 6, 2))
        SetOLEColour
    End If
    chkLowerCase_Click
End Sub

Private Sub txtOLEColour_Change()
    On Error GoTo txtOLEColourErr
    Dim OLEColour As Long
    OLEColour = Val(txtOLEColour.Text)
    If OLEColour >= 0 And OLEColour <= 16777215 And Not SetOLEColourExecuting Then
        txtRed.Text = OLEColour And &HFF&
        txtGreen.Text = (OLEColour And &HFF00&) \ &H100
        txtBlue.Text = (OLEColour And &HFF0000) \ &H10000
    End If
txtOLEColourErr:
End Sub

' Decimal red textbox change event handler
Private Sub txtRed_Change()
    On Error GoTo txtRedErr
    SliderR.Value = txtRed.Text
    If txtRedHex.Text <> Hex(txtRed.Text) Then txtRedHex.Text = Hex(txtRed.Text)
    If optBackground.Value Then
        lblMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    Else
        lblMain.ForeColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
    End If
txtRedErr:
End Sub

' Hexadecimal red textbox change event handler
Private Sub txtRedHex_Change()
    txtRed.Text = Dec(txtRedHex.Text)
    If txtRedHex.Text <> Mid(txtHTMLColour.Text, 2, 2) Then SetHTMLAndOLEColour
End Sub

' Converts a hexadecimal number from a string to a decimal number
Private Function Dec(hexNumber As String) As Long
    Dec = Val("&H" + hexNumber)
End Function

' Sets the HTML and OLE colour
Private Sub SetHTMLAndOLEColour()
    SetHTMLColour
    SetOLEColour
End Sub

' Sets the HTML colour
Private Sub SetHTMLColour()
    SetHTMLColourExecuting = True
    Dim R As String
    Dim G As String
    Dim B As String
    R = CStr(Hex(SliderR.Value))
    G = CStr(Hex(SliderG.Value))
    B = CStr(Hex(SliderB.Value))
    If Len(R) = 1 Then R = "0" + R
    If Len(G) = 1 Then G = "0" + G
    If Len(B) = 1 Then B = "0" + B
    txtHTMLColour.Text = "#" + R + G + B
    SetHTMLColourExecuting = False
End Sub

' Sets the OLE colour
Private Sub SetOLEColour()
    SetOLEColourExecuting = True
    txtOLEColour.Text = SliderR.Value + (SliderG.Value * 256) + (SliderB.Value * 256 * 256)
    SetOLEColourExecuting = False
End Sub
