VERSION 5.00
Begin VB.Form BMI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMI Calculator"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Izlaz"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdIzracunaj 
      Caption         =   "Izracunaj"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtTezina 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtVisina 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "BMI Calculator"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   3000
   End
   Begin VB.Label lblTezina 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vasa tezina:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "BMI Index:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   7
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label lblIndex 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tezina (kg)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Visina (cm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   990
   End
End
Attribute VB_Name = "BMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdIzracunaj_Click()
Dim a As Single, b As Single, c As Single

a = txtVisina.Text
b = txtTezina.Text

c = b / (a / 100) ^ 2
lblIndex.Caption = c

Select Case c
Case 0 To 14
    lblTezina.Caption = "Teska neuhranjenost"
Case 14 To 18
    lblTezina.Caption = "Neuhranjenost"
Case 18 To 21
    lblTezina.Caption = "Normalna tezina"
Case 21 To 23
    lblTezina.Caption = "Idealna tezina"
Case 23 To 25
    lblTezina.Caption = "Normalna tezina"
Case 25 To 30
    lblTezina.Caption = "Gojaznost"
Case 30 To 35
    lblTezina.Caption = "Teska gojaznost"
Case 35 < c
    lblTezina.Caption = "Opasna gojaznost"
End Select
End Sub

Private Sub cmdReset_Click()
txtVisina.Text = ""
txtTezina.Text = ""
lblIndex.Caption = ""
lblTezina.Caption = ""
End Sub

Private Sub Command1_Click()
End
End Sub
