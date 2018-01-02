VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmoffice 
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00B8F4FA&
      BorderStyle     =   0  'None
      Caption         =   "EMPLOYEE ID"
      Height          =   8895
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   8175
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
         Left            =   3240
         TabIndex        =   22
         Top             =   8280
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
         Left            =   960
         TabIndex        =   21
         Top             =   8280
         Width           =   1935
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "UPDATE"
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
         Left            =   5520
         TabIndex        =   5
         Top             =   7440
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
         Left            =   3240
         TabIndex        =   4
         Top             =   7440
         Width           =   1935
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
         Left            =   960
         TabIndex        =   3
         Top             =   7440
         Width           =   1935
      End
      Begin VB.TextBox txtsal 
         Height          =   375
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   2
         Top             =   5880
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmoffice.frx":0000
         Left            =   2880
         List            =   "frmoffice.frx":0010
         TabIndex        =   1
         Top             =   6600
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   41538
      End
      Begin VB.Label lblgender 
         BackStyle       =   0  'Transparent
         Caption         =   "GENDER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblmno 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label lbladdr 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lbldob 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF BIRTH"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblnm 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblid 
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYEE ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblsal 
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label lblpost 
         BackStyle       =   0  'Transparent
         Caption         =   "POST"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Label lblid1 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblnm1 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblgen 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lbladd 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2880
         TabIndex        =   8
         Top             =   3240
         Width           =   3855
      End
      Begin VB.Label lblmno1 
         BackColor       =   &H00E4D2BA&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   5040
         Width           =   3255
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
      Left            =   1680
      TabIndex        =   20
      Top             =   360
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   30000
      Left            =   0
      Picture         =   "frmoffice.frx":0058
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   45000
   End
End
Attribute VB_Name = "frmoffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim res As ADODB.Recordset

Private Sub cmddel_Click()
    Dim a As Integer

    a = InputBox("Enter The Agent ID to Delete : ")

    con.Execute ("Delete from Employee where E_id = '" & a & "'")
    con.Execute ("Delete from OfficeWorker where E_id = '" & a & "'")
    con.Execute ("commit")
    MsgBox (" Record Deleted Successfully")

End Sub

Private Sub cmdsubmit_Click()
    con.Execute ("insert into OfficeWorker(E_ID,E_JOBTITLE,E_SAL) values('" & lblid1.Caption & "','" & Combo1.Text & "','" & txtsal.Text & "')")
    con.Execute ("Commit")
    MsgBox ("Record Inserted Successfully............")
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    
    con.Open "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True"
    MsgBox ("Connection Established..........")
End Sub

