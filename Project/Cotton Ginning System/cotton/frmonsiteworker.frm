VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmonsiteworker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   7125
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdback 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   14
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   13
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   960
      TabIndex        =   10
      Top             =   240
      Width           =   10815
      Begin VB.Label lbladdg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "On the kanalda road ,near railway station, in shivaginagar , jalgaon"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -360
         TabIndex        =   12
         Top             =   1080
         Width           =   12855
      End
      Begin VB.Label lblginn 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LAKSHMI GINN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   9975
      End
      Begin VB.Image Image2 
         Height          =   9420
         Left            =   -600
         Picture         =   "frmonsiteworker.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   20880
      End
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   -1080
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Onsite Worker"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BFFBFA&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   1320
      TabIndex        =   1
      Top             =   2760
      Width           =   7455
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   41574
      End
      Begin VB.ComboBox cmb_eid 
         Height          =   315
         Left            =   2760
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox Comboshift 
         Height          =   315
         ItemData        =   "frmonsiteworker.frx":1B6EBA
         Left            =   2760
         List            =   "frmonsiteworker.frx":1B6EC7
         TabIndex        =   8
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtperdaywgs 
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtnohrs 
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lbldate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblperdaywgs 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PER DAY WEDGES"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblshift 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SHIFT"
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
         Left            =   0
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblnohrs 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NO OF HOURS"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblEid 
         Alignment       =   2  'Center
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
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   30000
      Left            =   -960
      Picture         =   "frmonsiteworker.frx":1B6ED4
      Top             =   0
      Width           =   45000
   End
End
Attribute VB_Name = "frmonsiteworker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim con As ADODB.Connection
Dim res As New ADODB.Recordset
Dim flag1 As Integer
Dim b As Integer

Private Sub CancelButton_Click()
Unload frmonsiteworker
End Sub

Private Sub cmb_eid_DropDown()
If (flag1 = 1) Then
 res.Open "select E_ID from Employee where E_DESIGNATION = 'Onsite Worker'", con, adOpenKeyset, adLockReadOnly, adCmdText
    While res.EOF <> True
         cmb_eid.AddItem (res(0))
    res.MoveNext
    Wend
res.Close
End If

flag1 = 2
End Sub

Private Sub cmdback_Click()
frmmain.Show
End Sub

Private Sub cmdclear_Click()
    cmb_eid.Text = ""
    txtnohrs.Text = ""
    Comboshift.Text = ""
    txtperdaywgs.Text = ""
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    
    con.Open "Provider=MSDAORA.1;Password=orant;User ID=system;Persist Security Info=True"
    MsgBox ("Connection Established..........")
    
    flag1 = 1
End Sub

Private Sub OKButton_Click()
    
    con.Execute ("insert into ONSITEWORKER (E_ID,E_NOHRS,E_SHIFT,E_WAGES,E_DAY,E_MONTH,E_YEAR) values ('" & cmb_eid.Text & "','" & txtnohrs.Text & "','" & Comboshift.Text & "','" & txtperdaywgs.Text & "','" & DTPicker1.Day & "','" & DTPicker1.Month & "','" & DTPicker1.Year & "')")
    MsgBox ("Record Inserted Successfullyyyyyy")
    
    cmb_eid.Text = ""
    txtnohrs.Text = ""
    Comboshift.Text = ""
    txtperdaywgs.Text = ""

End Sub

Private Sub txtperdaywgs_Click()

    b = Val(txtnohrs.Text)
    b = CInt(b * 100)
    txtperdaywgs.Text = b
End Sub
