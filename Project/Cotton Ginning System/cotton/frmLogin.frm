VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   8955
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5290.909
   ScaleMode       =   0  'User
   ScaleWidth      =   14422.2
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12720
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00601159&
      Caption         =   "LOGIN"
      Height          =   3375
      Left            =   6120
      TabIndex        =   0
      Top             =   4920
      Width           =   8175
      Begin VB.CommandButton cmdlogin 
         BackColor       =   &H00FFFFC0&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtpsw 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtuser 
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin VB.Image Image2 
         Height          =   2415
         Left            =   6000
         Picture         =   "frmLogin.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C63424&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   9351.896
      X2              =   9351.896
      Y1              =   2694.198
      Y2              =   1985.199
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C63424&
      BorderWidth     =   5
      X1              =   3492.877
      X2              =   9351.896
      Y1              =   1985.199
      Y2              =   1985.199
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C63424&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   3492.877
      X2              =   3492.877
      Y1              =   1276.199
      Y2              =   1985.199
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C63424&
      BorderWidth     =   5
      Height          =   1815
      Left            =   600
      Top             =   360
      Width           =   6855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C63424&
      BorderWidth     =   5
      Height          =   4095
      Left            =   5640
      Top             =   4560
      Width           =   9015
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   960
      Picture         =   "frmLogin.frx":5567
      Top             =   600
      Width           =   6210
   End
   Begin VB.Image Image3 
      Height          =   11520
      Left            =   0
      Picture         =   "frmLogin.frx":7312
      Top             =   -840
      Width           =   15360
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd2_Click()
Form1.Show
End Sub

Private Sub cmdlogin_Click()
    If txtuser.Text = "admin" And txtpsw.Text = "admin" Then
    MsgBox "LOGIN SUCCESSFULL."
        Unload frmLogin
        frmmain.Show
    Else
    MsgBox "PLEASE TRY AGAIN LATER"
    End If

End Sub

Private Sub Timer1_Timer()
If Shape1.BorderColor = &HC63424 Then
Shape1.BorderColor = &H0&
ElseIf Shape1.BorderColor = &H0& Then
Shape1.BorderColor = &HC63424
End If
If Shape2.BorderColor = &HC63424 Then
Shape2.BorderColor = &H0&
ElseIf Shape2.BorderColor = &H0& Then
Shape2.BorderColor = &HC63424
End If
If Line2.BorderColor = &HC63424 Then
Line2.BorderColor = &H0&
ElseIf Line2.BorderColor = &H0& Then
Line2.BorderColor = &HC63424
End If
If Line1.BorderColor = &HC63424 Then
Line1.BorderColor = &H0&
ElseIf Line1.BorderColor = &H0& Then
Line1.BorderColor = &HC63424
End If
If Line3.BorderColor = &HC63424 Then
Line3.BorderColor = &H0&
ElseIf Line3.BorderColor = &H0& Then
Line3.BorderColor = &HC63424
End If

End Sub
