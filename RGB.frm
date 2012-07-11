VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRGB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peter Chapman's Colour Slider 2005 (V2.5)"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "RGB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optForeground 
      Caption         =   "Foreground"
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   1680
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "Background"
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtHTMLColour 
      Height          =   285
      Left            =   4680
      MaxLength       =   7
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtGreen 
      Height          =   285
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   15
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtRedHex 
      Height          =   285
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   13
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
      TabIndex        =   11
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
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   4095
      TabIndex        =   6
      Top             =   2040
      Width           =   4095
   End
   Begin ComctlLib.Slider SliderB 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   327682
      Max             =   255
   End
   Begin ComctlLib.Slider SliderG 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   327682
      Max             =   255
   End
   Begin ComctlLib.Slider SliderR 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   327682
      Max             =   255
   End
   Begin VB.Label lblGeneral 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "HTML Colour"
      Height          =   195
      Index           =   5
      Left            =   3600
      TabIndex        =   17
      Top             =   1725
      Width           =   945
   End
   Begin VB.Label lblGeneral 
      Alignment       =   2  'Center
      Caption         =   "Hex"
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblGeneral 
      Alignment       =   2  'Center
      Caption         =   "Dec"
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Blue"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Green"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Red"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SetHTMLColourExecuting As Boolean

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtRed.Text = GetSetting("Peter Chapman", "Colour Slider", "Red", "0")
txtGreen.Text = GetSetting("Peter Chapman", "Colour Slider", "Green", "0")
txtBlue.Text = GetSetting("Peter Chapman", "Colour Slider", "Blue", "0")
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "Peter Chapman", "Colour Slider", "Red", txtRed.Text
SaveSetting "Peter Chapman", "Colour Slider", "Green", txtGreen.Text
SaveSetting "Peter Chapman", "Colour Slider", "Blue", txtBlue.Text
End Sub

Private Sub SliderB_Scroll()
txtBlue.Text = SliderB.Value
txtBlueHex.Text = Hex(SliderB.Value)
picMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
End Sub

Private Sub SliderG_Scroll()
txtGreen.Text = SliderG.Value
txtGreenHex.Text = Hex(SliderG.Value)
picMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
End Sub

Private Sub SliderR_Scroll()
txtRed.Text = SliderR.Value
txtRedHex.Text = Hex(SliderR.Value)
picMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
End Sub

Private Sub txtBlue_Change()
On Error GoTo txtBlueErr
SliderB.Value = txtBlue.Text
If txtBlueHex.Text <> Hex(txtBlue.Text) Then txtBlueHex.Text = Hex(txtBlue.Text)
picMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
txtBlueErr:
End Sub

Private Sub txtBlueHex_Change()
txtBlue.Text = Dec(txtBlueHex.Text)
If txtBlueHex.Text <> Mid(txtHTMLColour.Text, 6, 2) Then SetHTMLColour
End Sub

Private Sub txtGreen_Change()
On Error GoTo txtGreenErr
SliderG.Value = txtGreen.Text
If txtGreenHex.Text <> Hex(txtGreen.Text) Then txtGreenHex.Text = Hex(txtGreen.Text)
picMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
txtGreenErr:
End Sub

Private Sub txtGreenHex_Change()
txtGreen.Text = Dec(txtGreenHex.Text)
If txtGreenHex.Text <> Mid(txtHTMLColour.Text, 4, 2) Then SetHTMLColour
End Sub

Private Sub txtHTMLColour_Change()
Dim hexColour As String
If Left(txtHTMLColour.Text, 1) = "#" And SetHTMLColourExecuting = False Then
    hexColour = txtHTMLColour.Text
    If Len(hexColour) < 7 Then hexColour = hexColour + String(7 - Len(hexColour), "0")
    txtRedHex.Text = Mid(hexColour, 2, 2)
    txtGreenHex.Text = Mid(hexColour, 4, 2)
    txtBlueHex.Text = Mid(hexColour, 6, 2)
End If
End Sub

Private Sub txtRed_Change()
On Error GoTo txtRedErr
SliderR.Value = txtRed.Text
If txtRedHex.Text <> Hex(txtRed.Text) Then txtRedHex.Text = Hex(txtRed.Text)
picMain.BackColor = RGB(SliderR.Value, SliderG.Value, SliderB.Value)
txtRedErr:
End Sub

Private Sub txtRedHex_Change()
txtRed.Text = Dec(txtRedHex.Text)
If txtRedHex.Text <> Mid(txtHTMLColour.Text, 2, 2) Then SetHTMLColour
End Sub

Private Function Dec(hexNumber As String) As Long
Dec = Val("&H" + hexNumber)
End Function

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
