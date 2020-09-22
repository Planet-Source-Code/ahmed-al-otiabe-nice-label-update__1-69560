VERSION 5.00
Object = "*\AProject1.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dia 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   69891
   End
   Begin Project1.Label TEST 
      Height          =   1005
      Index           =   0
      Left            =   1155
      Top             =   60
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   1773
      Alignment       =   1
      BackStyle       =   9
      BorderColor1    =   -2147483647
      BorderColor2    =   -2147483643
      BroderStyle     =   4
      Caption         =   "TEST"
      Effects         =   4
      FillStyle       =   3
      EffectColor     =   -2147483627
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.Label TEST 
      Height          =   5000
      Index           =   1
      Left            =   60
      Top             =   60
      Width           =   1000
      _ExtentX        =   1773
      _ExtentY        =   8811
      Alignment       =   1
      BackStyle       =   9
      BorderColor1    =   -2147483647
      BorderColor2    =   -2147483643
      BroderStyle     =   4
      Caption         =   "TEST"
      Effects         =   4
      FillStyle       =   3
      EffectColor     =   -2147483627
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Rotation        =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   1150
      TabIndex        =   0
      Top             =   1100
      Width           =   5000
      Begin VB.ComboBox ThemeColor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":6852
         Left            =   2600
         List            =   "Form1.frx":6868
         TabIndex        =   16
         Text            =   "NoTheme"
         Top             =   480
         Width           =   2300
      End
      Begin VB.ComboBox BackStyle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":68A8
         Left            =   120
         List            =   "Form1.frx":68CA
         TabIndex        =   15
         Text            =   "BackFill_VerticalOut"
         Top             =   1080
         Width           =   2300
      End
      Begin VB.PictureBox Picture1 
         Height          =   250
         Left            =   1600
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   14
         Top             =   1440
         Width           =   500
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000005&
         Height          =   250
         Left            =   1600
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   13
         Top             =   1800
         Width           =   500
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000001&
         Height          =   250
         Left            =   4100
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   12
         Top             =   1440
         Width           =   500
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000015&
         Height          =   250
         Left            =   4100
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   11
         Top             =   2640
         Width           =   500
      End
      Begin VB.ComboBox Effects 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":69A4
         Left            =   2600
         List            =   "Form1.frx":69C6
         TabIndex        =   10
         Text            =   "Shadow"
         Top             =   2280
         Width           =   2300
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000005&
         Height          =   250
         Left            =   4100
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   9
         Top             =   1755
         Width           =   500
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FF0000&
         Height          =   250
         Left            =   1600
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   8
         Top             =   3000
         Width           =   500
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   250
         Left            =   1600
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   7
         Top             =   2640
         Width           =   500
      End
      Begin VB.ComboBox mFillStyle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":6A34
         Left            =   120
         List            =   "Form1.frx":6A53
         TabIndex        =   6
         Text            =   "TextFill_VerticalDown"
         Top             =   2280
         Width           =   2300
      End
      Begin VB.ComboBox Alignment 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":6B20
         Left            =   120
         List            =   "Form1.frx":6B2D
         TabIndex        =   5
         Text            =   "Center_Justify"
         Top             =   480
         Width           =   2300
      End
      Begin VB.ComboBox BroderStyle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":6B5E
         Left            =   2600
         List            =   "Form1.frx":6B89
         TabIndex        =   4
         Text            =   "Broder_3DInside"
         Top             =   1080
         Width           =   2300
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Text            =   "TEST"
         Top             =   3500
         Width           =   2800
      End
      Begin VB.CommandButton Command2 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4250
         TabIndex        =   2
         Top             =   3400
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3000
         TabIndex        =   1
         Top             =   3400
         Width           =   1200
      End
      Begin Project1.Label Label9 
         Height          =   195
         Left            =   2600
         Top             =   840
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "BorderStyle"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label1 
         Height          =   200
         Left            =   120
         Top             =   240
         Width           =   1250
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   " Alignment"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label13 
         Height          =   195
         Left            =   2595
         Top             =   2640
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "EffectColor >> "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label10 
         Height          =   195
         Left            =   2595
         Top             =   2040
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "Effects"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label11 
         Height          =   195
         Left            =   2600
         Top             =   1440
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "BroderColor1 >> "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label12 
         Height          =   195
         Left            =   2600
         Top             =   1755
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "BroderColor2 >> "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label7 
         Height          =   195
         Left            =   120
         Top             =   3000
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "ForeColor2 >> "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label5 
         Height          =   195
         Left            =   120
         Top             =   2700
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "ForeColor1 >> "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label3 
         Height          =   195
         Left            =   120
         Top             =   2040
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "FillStyle"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label LabelText 
         Height          =   375
         Left            =   75
         Top             =   3450
         Width           =   2875
         _ExtentX        =   5080
         _ExtentY        =   661
         BorderColor1    =   -2147483647
         BorderColor2    =   16777215
         BroderStyle     =   11
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label8 
         Height          =   195
         Left            =   120
         Top             =   1800
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "BackColor2 >> "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label6 
         Height          =   195
         Left            =   120
         Top             =   1500
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "BackColor1 >> "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label4 
         Height          =   195
         Left            =   120
         Top             =   840
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "BackStyle"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.Label Label2 
         Height          =   195
         Left            =   2600
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         BackStyle       =   0
         BorderColor1    =   -2147483647
         BorderColor2    =   -2147483643
         Caption         =   "ThemeColor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim R           As Integer
Dim Warning     As Boolean
Private Sub Form_Load()
    '>> Form1 Centered
    Form1.Move (Screen.Width - Form1.Width) / 2, (Screen.Height - Form1.Height) / 2
    '>> Form1 Size's
    Form1.Width = 6300
    Form1.Height = 5600
End Sub
Private Sub Alignment_Click()
    For R = 0 To 1
        TEST(R).Alignment = Alignment.ListIndex
    Next R
End Sub
Private Sub Picture7_Click()
    Dia.CancelError = False
    Dia.ShowColor
    Picture7.BackColor = Dia.Color
    For R = 0 To 1
        TEST(R).EffectColor = Picture7.BackColor
    Next R
    Picture7.BackColor = TEST(0).EffectColor
End Sub
Private Sub ThemeColor_Click()
    For R = 0 To 1
        TEST(R).ThemeColor = ThemeColor.ListIndex
    Next R
End Sub
Private Sub mFillStyle_Click()
    For R = 0 To 1
        TEST(R).FillStyle = mFillStyle.ListIndex
    Next R
End Sub
Private Sub BackStyle_Click()
    For R = 0 To 1
        TEST(R).BackStyle = BackStyle.ListIndex
    Next R
End Sub
Private Sub BroderStyle_Click()
    For R = 0 To 1
        TEST(R).BroderStyle = BroderStyle.ListIndex
    Next R
End Sub
Private Sub Effects_Click()
    For R = 0 To 1
        TEST(R).Effects = Effects.ListIndex
    Next R
        Picture7.BackColor = TEST(0).EffectColor
End Sub
Private Sub Picture1_Click()
    Dia.CancelError = False
    Dia.ShowColor
    Picture1.BackColor = Dia.Color
    For R = 0 To 1
        TEST(R).BackColor1 = Picture1.BackColor
    Next R
    Picture1.BackColor = TEST(0).BackColor1
End Sub
Private Sub Picture2_Click()
    Dia.CancelError = False
    Dia.ShowColor
    Picture2.BackColor = Dia.Color
    For R = 0 To 1
        TEST(R).BackColor2 = Picture2.BackColor
    Next R
    Picture2.BackColor = TEST(1).BackColor2
End Sub
Private Sub Picture3_Click()
    Dia.CancelError = False
    Dia.ShowColor
    Picture3.BackColor = Dia.Color
    For R = 0 To 1
        TEST(R).ForeColor1 = Picture3.BackColor
    Next R
    Picture3.BackColor = TEST(0).ForeColor1
End Sub
Private Sub Picture4_Click()
    Dia.CancelError = False
    Dia.ShowColor
    Picture4.BackColor = Dia.Color
    For R = 0 To 1
        TEST(R).ForeColor2 = Picture4.BackColor
    Next R
    Picture4.BackColor = TEST(1).ForeColor2
End Sub
Private Sub Picture5_Click()
    Dia.CancelError = False
    Dia.ShowColor
    Picture5.BackColor = Dia.Color
    For R = 0 To 1
        TEST(R).BorderColor1 = Picture5.BackColor
    Next R
    Picture5.BackColor = TEST(0).BorderColor1
End Sub
Private Sub Picture6_Click()
    Dia.CancelError = False
    Dia.ShowColor
    Picture6.BackColor = Dia.Color
    For R = 0 To 1
        TEST(R).BorderColor2 = Picture6.BackColor
    Next R
    Picture6.BackColor = TEST(1).BorderColor2
End Sub
Private Sub Text1_Change()
    For R = 0 To 1
        TEST(R).Caption = Text1.Text
    Next R
End Sub
Private Sub Command1_Click()
On Error GoTo ErrHandler
   If Warning = False Then
    '>> My Properties.
    Dia.FontName = "Times New Roman"
    Dia.FontSize = 48
    Dia.FontBold = True
    Dia.FontItalic = True
      MsgBox "In Case Of The Label Is (Ratation = Vertical) Please Use True Type Font Only.", vbExclamation, "Warning"
      Warning = True
   End If
    '......................................
    '>> Open Common Dialog Font.
    Dia.ShowFont
    Font.Name = Dia.FontName
    Font.Bold = Dia.FontBold
    Font.Italic = Dia.FontItalic
    Font.Underline = Dia.FontUnderline
    Font.Size = Dia.FontSize
    For R = 0 To 1
        Set TEST(R).Font = Font
    Next R
ErrHandler:
End Sub
Private Sub Command2_Click()
    TEST(1).ShowAbout
End Sub

