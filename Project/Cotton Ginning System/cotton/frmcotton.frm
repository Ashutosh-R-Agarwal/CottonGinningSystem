VERSION 5.00
Begin VB.Form frmcotton 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   Picture         =   "frmcotton.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmr1 
      Interval        =   500
      Left            =   1200
      Top             =   1800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0026060E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9375
      Left            =   2280
      TabIndex        =   0
      Top             =   1200
      Width           =   15975
      Begin VB.CommandButton cmdsubmit 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1200
         TabIndex        =   15
         Top             =   7920
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E4D2BA&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6255
         Left            =   1080
         TabIndex        =   2
         Top             =   1200
         Width           =   13695
         Begin VB.TextBox txtagent 
            Height          =   615
            Left            =   4800
            TabIndex        =   14
            Top             =   4320
            Width           =   4215
         End
         Begin VB.TextBox txtrate 
            Height          =   615
            Left            =   4800
            TabIndex        =   13
            Top             =   3360
            Width           =   4215
         End
         Begin VB.TextBox txtqty 
            Height          =   615
            Left            =   4800
            TabIndex        =   12
            Top             =   2400
            Width           =   4215
         End
         Begin VB.TextBox txtvar 
            Height          =   615
            Left            =   4800
            TabIndex        =   11
            Top             =   1560
            Width           =   4215
         End
         Begin VB.TextBox txtid 
            Height          =   615
            Left            =   4800
            TabIndex        =   10
            Top             =   720
            Width           =   4215
         End
         Begin VB.Label lblagent 
            BackStyle       =   0  'Transparent
            Caption         =   "AGENT NAME"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   9
            Top             =   4440
            Width           =   1935
         End
         Begin VB.Label lblrate 
            BackStyle       =   0  'Transparent
            Caption         =   "RATE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   8
            Top             =   3480
            Width           =   1935
         End
         Begin VB.Label lblqty 
            BackStyle       =   0  'Transparent
            Caption         =   "QUANTITY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   7
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label lblvar 
            BackStyle       =   0  'Transparent
            Caption         =   "VARIETY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   6
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lblid 
            BackStyle       =   0  'Transparent
            Caption         =   "COTTON ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   5
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdprevious 
         Caption         =   "PREVIOUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   12000
         TabIndex        =   4
         Top             =   7920
         Width           =   2655
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6600
         TabIndex        =   3
         Top             =   7920
         Width           =   2655
      End
   End
   Begin VB.Label lblcotton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COTTON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   7680
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   14295
      Left            =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "frmcotton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsubmit_Click()
If txtid.Text = "" Or txtvar.Text = "" Or txtqty.Text = "" Or txtrate.Text = "" Or txtagent.Text = "" Then
MsgBox "PLEASE FILL THE DETAILS"
End If
frmexport.Show
End Sub

Private Sub Label1_Click()

End Sub

Private Sub tmr1_Timer()
If lblcotton.ForeColor = &H800000 Then
lblcotton.ForeColor = &HC00000
ElseIf lblcotton.ForeColor = &HC00000 Then
lblcotton.ForeColor = &H800000
End If

End Sub

Private Sub txtagent_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case vbKey0 To vbKey9
KeyAscii = 0
Beep
End Select
End Sub
