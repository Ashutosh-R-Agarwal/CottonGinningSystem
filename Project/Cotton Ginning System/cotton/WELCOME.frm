VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmwelcome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   Picture         =   "WELCOME.frx":0000
   ScaleHeight     =   7020
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   5280
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   7
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1920
      Top             =   4920
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading........"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   4200
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait....."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   3120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COTTON GINNING SYSTEM"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   6735
   End
End
Attribute VB_Name = "frmwelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Enabled = False
    ProgressBar1.Value = 0
    ProgressBar1.Value = 0
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    r = Rnd() * 255
    g = Rnd() * 255
    b = Rnd() * 255
    
    Label1.ForeColor = RGB(g, r, b)
    Label2.ForeColor = RGB(g, r, b)
    Label3.ForeColor = RGB(g, r, b)
    
    ProgressBar1.Value = ProgressBar1.Value + 1
    
    If ProgressBar1.Value = 7 Then Timer1.Enabled = False
    If Timer1.Enabled = False Then
    Unload frmwelcome
    frmLogin.Show
'    Me.Hide

End If

End Sub
