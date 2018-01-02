VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmagent 
   Caption         =   "Form2"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   10455
   Begin MSAdodcLib.Adodc AdodcAgentfrm 
      Height          =   615
      Left            =   8760
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "Provider=MSDAORA.1;Password=aditya;User ID=system;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=aditya;User ID=system;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "AGENT"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Caption         =   "agent details"
      Height          =   8535
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8295
      Begin VB.TextBox txtdate 
         Height          =   285
         Left            =   4440
         TabIndex        =   19
         ToolTipText     =   "eg : 09-Jan-1990"
         Top             =   5640
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "AGENT.frx":0000
         Left            =   2280
         List            =   "AGENT.frx":000A
         TabIndex        =   18
         Top             =   4920
         Width           =   1575
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
         Height          =   495
         Left            =   3480
         TabIndex        =   17
         Top             =   7440
         Width           =   1335
      End
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
         Height          =   495
         Left            =   1800
         TabIndex        =   16
         Top             =   7440
         Width           =   1455
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
         Height          =   495
         Left            =   5040
         TabIndex        =   14
         Top             =   6840
         Width           =   1455
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
         Height          =   495
         Left            =   3480
         TabIndex        =   13
         Top             =   6840
         Width           =   1335
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
         Height          =   495
         Left            =   1800
         TabIndex        =   12
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox txtaddr 
         Height          =   1215
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3240
         Width           =   4095
      End
      Begin VB.TextBox txtmno 
         Height          =   375
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtnm 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox txtid 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   1320
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   5640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   145227777
         CurrentDate     =   41538
      End
      Begin VB.Label lblagent 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   15
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label lbladdr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblmno 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblnm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblid 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Agent ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblgender 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label lbldob 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   5640
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmagent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Command

Private Sub cmdsubmit_Click()
frmcotton.Show
End Sub

Private Sub cmdAdd_Click()
    con.Execute ("insert into AGENT values('" & txtid.Text & "','" & txtnm.Text & "','" & txtdate.Text & "','" & Combo1.Text & "','" & txtaddr.Text & "','" & txtmno.Text & "')")
    con.Execute ("commit")
    MsgBox ("Record Successfully Inserted.......")
End Sub

Private Sub cmdClear_Click()
    txtid.Text = ""
    txtnm.Text = ""
    txtmno.Text = ""
    txtdate.Text = ""
    txtaddr.Text = ""
    Combo1.Text = ""
End Sub

Private Sub cmdUpdate_Click()
Dim a As Integer

a = InputBox("Enter The Agent ID to Update : ")

con.Execute ("update AGENT set where AID = 'a'")
txtid.Text = res(0)
txtnm.Text = res(1)
txtdate.Text = res(2)
Combo1.Text = res(3)
txtaddr.Text = res(4)
txtmno.Text = res(5)

MsgBox ("Record Updated Successfully.....")
End Sub

Private Sub Form_Load()
    
    Set con = New ADODB.Connection
    Set res = New ADODB.Command
    
    con.Open "Provider=MSDAORA.1;Password=aditya;User ID=system;Persist Security Info=True"
    MsgBox ("Connection Established..........")
End Sub

Private Sub Form_Resize()
    frmagent.Width = 9400
    frmagent.Height = 9800
End Sub

Private Sub tmr1_Timer()
If lblagent.ForeColor = &HFFC0C0 Then
lblagent.ForeColor = &HC00000
ElseIf lblagent.ForeColor = &HC00000 Then
lblagent.ForeColor = &HFFC0C0
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
