VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmexport 
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19095
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "frmexport.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   19095
   WindowState     =   2  'Maximized
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
      Left            =   16080
      TabIndex        =   31
      Top             =   9600
      Width           =   1935
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
      Left            =   18240
      TabIndex        =   30
      Top             =   9600
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   960
      TabIndex        =   27
      Top             =   240
      Width           =   18495
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
         Left            =   3480
         TabIndex        =   29
         Top             =   1200
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
         Left            =   4920
         TabIndex        =   28
         Top             =   360
         Width           =   9975
      End
      Begin VB.Image Image1 
         Height          =   9420
         Left            =   -1320
         Picture         =   "frmexport.frx":159D6B
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   20880
      End
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   26
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton cmdclr 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11640
      TabIndex        =   25
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   24
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton cmdupdt 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   23
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   22
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "VIEW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13800
      TabIndex        =   21
      Top             =   9600
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   840
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
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
      RecordSource    =   "Company"
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONTACT PERSON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4815
      Left            =   9840
      TabIndex        =   10
      Top             =   4440
      Width           =   8055
      Begin VB.TextBox txtcpphno 
         Height          =   495
         Left            =   2520
         TabIndex        =   20
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox txtcpadd 
         Height          =   1215
         Left            =   2520
         ScrollBars      =   3  'Both
         TabIndex        =   19
         Top             =   1800
         Width           =   3615
      End
      Begin VB.ComboBox cmbgender 
         Height          =   315
         ItemData        =   "frmexport.frx":310C25
         Left            =   2520
         List            =   "frmexport.frx":310C2F
         TabIndex        =   18
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtcpemail 
         Height          =   495
         Left            =   2520
         TabIndex        =   16
         Top             =   3240
         Width           =   3615
      End
      Begin VB.TextBox txtcpname 
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblcpgender 
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
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL ID"
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
         TabIndex        =   15
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO"
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
         TabIndex        =   14
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   13
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COMPANY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   1440
      TabIndex        =   1
      Top             =   4440
      Width           =   8055
      Begin VB.TextBox txtadd 
         Height          =   1215
         Left            =   3240
         TabIndex        =   5
         Top             =   3000
         Width           =   3855
      End
      Begin VB.ComboBox cmbtype 
         Height          =   315
         ItemData        =   "frmexport.frx":310C39
         Left            =   3240
         List            =   "frmexport.frx":310C43
         TabIndex        =   4
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtpno 
         Height          =   315
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox txtnm 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label lbladd 
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY ADDRESS"
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
         Left            =   600
         TabIndex        =   9
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label lblcp 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO"
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
         Left            =   720
         TabIndex        =   8
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lbltype 
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
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
         Left            =   720
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY NAME"
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
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Timer tmr1 
      Interval        =   500
      Left            =   240
      Top             =   9120
   End
   Begin VB.Label lblexport 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   3240
      Width           =   5775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmexport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset
Dim up As String

Private Sub cmdsubmit_Click()
    If txtid.Text = "" Or txtnm.Text = "" Or txtstk.Text = "" Or txtmgr.Text = "" Then
        MsgBox "PLEASE FILL THE DETAILS"
    End If
    frmprocess.Show
End Sub


Private Sub cmdadd_Click()
If txtnm.Text = "" Or cmbtype.Text = "" Or txtpno.Text = "" Or txtadd.Text = "" Or txtcpname.Text = "" Or cmbgender.Text = "" Or txtcpadd.Text = "" Or txtcpemail.Text = "" Or txtcpphno.Text = "" Then
    MsgBox ("PLEASE FILL THE DETAILS")
    Else
    Dim a As Integer
    
    con.Execute ("insert into company(com_name,com_phno,com_add,com_type) values('" & txtnm.Text & "','" & txtpno.Text & "','" & txtadd.Text & "','" & cmbtype.Text & "')")
    
    con.Execute ("commit")
    res.Open "SELECT * from Company where Com_phno = '" & txtpno.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    a = res(0)
    con.Execute ("insert into contactperson(cp_name,cp_gender,cp_email,cp_phno,com_id) values('" & txtcpname.Text & "','" & cmbgender.Text & " ','" & txtcpemail.Text & "','" & txtcpphno.Text & "','" & a & "')")
    MsgBox ("YOUR COMPANY ID IS:" & a)
    MsgBox ("REcord Inserted Successfully........")
    res.Close
    End If
End Sub

Private Sub cmdback_Click()
Unload frmexport
frmmain.Show
End Sub

Private Sub cmdclr_Click()
    txtnm.Text = ""
    txtpno.Text = ""
    txtadd.Text = ""
    cmbtype.Text = ""
    
    txtcpname.Text = ""
    cmbgender.Text = ""
    txtcpadd.Text = ""
    txtcpemail.Text = ""
    txtcpphno.Text = ""
    
End Sub

Private Sub cmddel_Click()
    Dim a As String
    
    a = InputBox("Enter The Company ID to Delete : ", vbOKCancel)
    If a <> "" Then
    
    res.Open "SELECT * from Company where Com_id = '" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    If res.BOF <> True Then
    con.Execute ("delete from ContactPerson where Com_id ='" & a & "'")
        con.Execute ("delete from Company where com_id ='" & a & "'")
        
        con.Execute ("COMMIT")
        MsgBox ("Record Deleted Successfully..........")
    Else
        MsgBox ("RECORD NOT FOUND........")
    End If
    
    res.Close
    End If
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdsearch_Click()
    
    txtnm.Text = ""
    txtpno.Text = ""
    txtadd.Text = ""
    cmbtype.Text = ""
    
    txtcpname.Text = ""
    cmbgender.Text = ""
    txtcpadd.Text = ""
    txtcpemail.Text = ""
    txtcpphno.Text = ""
    
    
    'Dim s As String
    'Dim a As Integer
    
    up = InputBox("Enter The Company ID to Update : ", vbOKCancel)
    If up <> "" Then
    
    res.Open "SELECT * from Company where Com_id = '" & up & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    res1.Open "SELECT * from ContactPerson where Com_id = '" & up & "'", con, adOpenKeyset, adLockReadOnly, adCmdText

    If res.BOF <> True Then
        
        txtnm.Text = res(1)
        txtpno.Text = res(2)
        txtadd.Text = res(3)
        cmbtype.Text = res(4)
        
        txtcpname.Text = res1(0)
        cmbgender.Text = res1(1)
        txtcpadd.Text = res1(2)
        txtcpemail.Text = res1(3)
        txtcpphno.Text = res1(4)
        MsgBox ("RECORD Found SUCCESSFULLY........")
        
        res.MoveFirst
        res1.MoveFirst

        
    Else
    
        MsgBox ("RECORD NOT FOUND........")
    
    End If

    res.Close
    res1.Close
    End If
    
End Sub

Private Sub cmdupdt_Click()
    'Dim s As String
    'Dim a As Integer
    
    'a = InputBox("Enter The Company ID to Update : ")

    'res.Open "SELECT * from Company where Com_id = '" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    'res1.Open "SELECT * from ContactPerson where Com_id = '" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText

    'If res.BOF <> True Then
        con.Execute ("update company set com_name='" & txtnm.Text & "',com_phno='" & txtpno.Text & "',com_add='" & txtadd.Text & "',com_type='" & cmbtype.Text & "' where com_id= '" & up & "'")
        con.Execute ("commit")
        'res.Close
        'res1.Close
        'res.Open "SELECT * from Company where Com_phno = '" & txtpno.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
        'a = res(0)
        con.Execute ("update contactperson set cp_name='" & txtcpname.Text & "',cp_gender='" & cmbgender.Text & "',cp_add='" & txtadd.Text & "',cp_email='" & txtcpemail.Text & "',cp_phno='" & txtcpphno.Text & "',com_id='" & up & "' where com_id= '" & a & "'")
        MsgBox ("RECORD Updated SUCCESSFULLY........")

    'Else
        'MsgBox ("RECORD NOT FOUND........")
    'End If

End Sub


Private Sub cmdview_Click()
frmcomdetails.Show
End Sub


Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    
    con.Open "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True"
    MsgBox ("Connection Established..........")
End Sub

Private Sub tmr1_Timer()
    If lblexport.ForeColor = &H400040 Then
        lblexport.ForeColor = &HC00000
    ElseIf lblexport.ForeColor = &HC00000 Then
        lblexport.ForeColor = &H400040
    End If

End Sub

Private Sub txtcpemail_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii = Asc("@") Or KeyAscii = Asc("_") Or KeyAscii = Asc(".") Then
Else
    MsgBox ("INVALID INPUT")
End If

End Sub

Private Sub txtcpphno_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtnm_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        KeyAscii = 0
        Beep
    End Select

End Sub

Private Sub txtpno_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub
