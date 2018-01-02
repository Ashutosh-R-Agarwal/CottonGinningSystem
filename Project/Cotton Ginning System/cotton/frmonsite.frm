VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmonsite 
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmonsite.frx":0000
   ScaleHeight     =   9570
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00A3F1EF&
      BorderStyle     =   0  'None
      Caption         =   "EMPLOYEE ID"
      Height          =   8415
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   8775
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   24
         Top             =   7680
         Width           =   1935
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   7680
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmonsite.frx":159D6B
         Left            =   6120
         List            =   "frmonsite.frx":159D78
         TabIndex        =   6
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   5
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton cmdsubmit 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   6960
         Width           =   1935
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   6960
         Width           =   1935
      End
      Begin VB.CommandButton cmdprevious 
         Caption         =   "PREVIOUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   2
         Top             =   6960
         Width           =   1935
      End
      Begin VB.TextBox txtwag 
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   6360
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   41538
      End
      Begin VB.Label lblmno1 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   4680
         Width           =   3255
      End
      Begin VB.Label lbladd 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3000
         TabIndex        =   20
         Top             =   3120
         Width           =   3855
      End
      Begin VB.Label lblgen 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblnm1 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblid1 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblshift 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label lblno 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No Of Hours Worked"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   15
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label lblid 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblnm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lbldob 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lbladdr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblmno 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label lblgender 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblweg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Per Day Wages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   6360
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE DETAILS"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   1560
      TabIndex        =   22
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frmonsite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim con As ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset

Private Sub cmdsubmit_Click()
    
    Dim b As Integer
    
    If (Combo1.Text = "First Shift") Then
        b = 1
    ElseIf (Combo1.Text = "Second Shift") Then
        b = 2
    Else
        b = 3
    End If
    
    con.Execute ("insert into OnsiteWorker(E_ID,E_nohrs,E_shift,E_wages) values('" & lblid1.Caption & "','" & txtno.Text & "','" & b & "','" & txtwag.Text & "')")
    con.Execute ("Commit")
    MsgBox ("Record Inserted Successfully............")

End Sub

Private Sub Form_Load()

    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    
    con.Open "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True"
    MsgBox ("Connection Established..........")
    
    a = InputBox("Enter The no of hours Worked")
    txtno.Text = a
    txtwag.Text = CInt(a) * 100

End Sub
 
