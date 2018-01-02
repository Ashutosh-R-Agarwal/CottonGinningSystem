VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmagent 
   Caption         =   "Form2"
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Z.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
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
      Height          =   1095
      Left            =   12720
      TabIndex        =   25
      Top             =   8280
      Width           =   2535
   End
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
      Height          =   1095
      Left            =   9840
      TabIndex        =   24
      Top             =   8280
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   960
      TabIndex        =   21
      Top             =   120
      Width           =   18495
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
         Left            =   5160
         TabIndex        =   23
         Top             =   360
         Width           =   9975
      End
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
         TabIndex        =   22
         Top             =   1200
         Width           =   12855
      End
      Begin VB.Image Image2 
         Height          =   9420
         Left            =   -720
         Picture         =   "Z.frx":159D6B
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   20880
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   8535
      Left            =   9480
      TabIndex        =   14
      Top             =   840
      Width           =   6135
      Begin VB.CommandButton cmdSearch 
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
         Height          =   1095
         Left            =   360
         TabIndex        =   20
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton cmdClear 
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
         Height          =   1095
         Left            =   360
         TabIndex        =   19
         Top             =   6120
         Width           =   2535
      End
      Begin VB.CommandButton cmdAdd 
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
         Height          =   1095
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CommandButton cmdview 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3240
         TabIndex        =   17
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CommandButton cmdUpdate 
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
         Height          =   1095
         Left            =   3240
         TabIndex        =   16
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton cmdDelete 
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
         Height          =   1095
         Left            =   3240
         TabIndex        =   15
         Top             =   6120
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   30000
         Left            =   -360
         Picture         =   "Z.frx":310C25
         Top             =   -240
         Width           =   45000
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   720
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
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
      RecordSource    =   "Agent"
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
   Begin VB.Timer Timer1 
      Left            =   2880
      Top             =   10200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "agent details"
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   8295
      Begin VB.TextBox txtemail 
         Height          =   495
         Left            =   2160
         TabIndex        =   13
         ToolTipText     =   "abc.km@gmail.com"
         Top             =   6240
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Z.frx":46A990
         Left            =   2160
         List            =   "Z.frx":46A99A
         TabIndex        =   11
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox txtaddr 
         Height          =   1215
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   3240
         Width           =   4095
      End
      Begin VB.TextBox txtmno 
         Height          =   375
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtnm 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   1440
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   5640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   65929217
         CurrentDate     =   41538
      End
      Begin VB.Label lblemail 
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label lblagent 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "AGENT DETAILS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   5295
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblmno 
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   1575
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   4920
         Width           =   1575
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
         ForeColor       =   &H00000040&
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   5640
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmagent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset
Dim up As String

Private Sub cmdsubmit_Click()
frmcotton.Show
End Sub

Private Sub cmdadd_Click()
If txtnm.Text = "" Or txtmno.Text = "" Or txtaddr.Text = "" Or Combo1.Text = "" Or txtemail.Text = "" Then
MsgBox ("Please Fill The Details")
Else

    
    Dim a As Integer
    Dim s As String
    Dim d As String
    Dim Y As String
    Dim g As Integer
    
    'If (txtnm.Text = " " & txtmno.Text = " " & txtaddr.Text = " " & Combo1.Text = " " & txtemail.Text = " ") Then
     '   MsgBox ("Please Fill The Details Properly....")
    'Else
    
    s = Month(DTPicker1.Month)
    con.Execute ("insert into AGENT (A_name,A_gender,A_add,A_email,A_phno,DAY,MONTH,YEAR) values('" & txtnm.Text & "','" & Combo1.Text & "','" & txtaddr.Text & "','" & txtemail.Text & "','" & txtmno.Text & "','" & DTPicker1.Day & "','" & DTPicker1.Month & "','" & DTPicker1.Year & "')")
    res1.Open "SELECT A_id from Agent where A_name='" & txtnm.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    a = Val(res1(0))
    
    MsgBox ("YOUR AGENT ID IS: " & a)
    txtnm.Text = ""
    txtmno.Text = ""

    txtaddr.Text = ""
    Combo1.Text = ""
    txtemail.Text = ""
    res1.Close
   'End If
    
   MsgBox ("Record Successfully Inserted.......")
   End If
End Sub

Private Sub cmdback_Click()
Unload frmagent
frmmain.Show
End Sub

Private Sub cmdclear_Click()
    
    txtnm.Text = ""
    txtmno.Text = ""

    txtaddr.Text = ""
    Combo1.Text = ""
    txtemail.Text = ""
End Sub

Private Sub cmdDelete_Click()
Dim a As String
Dim b As String

a = InputBox("Enter The Agent ID to Delete : ", vbOKCancel)
If a <> "" Then
res.Open "SELECT A_ID from DEPENDS where A_id = '" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
If (res.EOF = True) Then


res.Close

res.Open "SELECT * from Agent where A_id = '" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText

If res.BOF <> True Then
con.Execute ("delete from Agent where A_id ='" & a & "'")
con.Execute ("COMMIT")
MsgBox ("Record Deleted Successfully..........")
Else
MsgBox ("RECORD NOT FOUND........")
End If
res.Close
Else
res.Close

MsgBox "CANNOT DELETE AGENT FROM DATABASE SINCE HE IS ALREADY INVOLVED IN A TRANSACTION"
End If
End If

End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdsearch_Click()

    
    txtnm.Text = ""
    txtmno.Text = ""
    
    txtaddr.Text = ""
    Combo1.Text = ""
    txtemail.Text = ""
    txtmno.Text = ""
    Dim s As String
    'Dim a As String
    
up = InputBox("Enter The Agent ID to Search : ", vbOKCancel)

If up <> "" Then
res.Open "SELECT * from Agent where A_id = '" & up & "'", con, adOpenKeyset, adLockReadOnly, adCmdText

If res.BOF <> True Then
txtnm.Text = res(1)
DTPicker1.Day = res(6)
DTPicker1.Month = res(7)
DTPicker1.Year = res(8)
Combo1.Text = res(2)
txtaddr.Text = res(3)
txtemail.Text = res(4)
txtmno.Text = res(5)

MsgBox ("RECORD FOUND SUCCESSFULLY........")

Else
MsgBox ("RECORD NOT FOUND........")
End If
res.Close
End If
End Sub

Private Sub cmdupdate_Click()




    
'a = InputBox("Enter The Agent ID to Update : ", vbOKCancel)

'If a <> "" Then
'res.Open "SELECT * from Agent where A_id = '" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText

'If res.BOF <> True Then

con.Execute ("update AGENT set A_name='" & txtnm.Text & "', A_gender = '" & Combo1.Text & "' ,A_add='" & txtaddr.Text & "',A_email='" & txtemail.Text & "',DAY='" & DTPicker1.Day & "',month='" & DTPicker1.Month & "',year='" & DTPicker1.Year & "' where A_id='" & up & "'")
'con.Execute ("update AGENT set A_name='" & txtnm.Text & "',A_dob='" & DTPicker1.Value & "',A_gender='" & Combo1.Text & "',A_add='" & txtaddr.Text & "',A_email='" & txtemail.Text & "' where A_id='" & a & "'")

con.Execute ("commit")
MsgBox ("Record Updated Successfully.....")


'Else
'MsgBox ("RECORD NOT FOUND........")
'End If
'res.Close
'End If
End Sub

Private Sub cmdview_Click()
    frmagdetails.Show
End Sub


Private Sub Form_Load()

    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    
    con.Open "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True"
    MsgBox ("Connection Established..........")
End Sub

Private Sub tmr1_Timer()
If lblagent.ForeColor = &HFFC0C0 Then
lblagent.ForeColor = &HC00000
ElseIf lblagent.ForeColor = &HC00000 Then
lblagent.ForeColor = &HFFC0C0
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    con.Close
    MsgBox "Connection Closed............"
End Sub

Private Sub txtid_Change()
txtid.Text = agent_seq
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii = Asc("@") Or KeyAscii = Asc("_") Or KeyAscii = Asc(".") Then
Else
    MsgBox ("INVALID INPU")
End If
End Sub

Private Sub txtmno_KeyPress(KeyAscii As Integer)
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
