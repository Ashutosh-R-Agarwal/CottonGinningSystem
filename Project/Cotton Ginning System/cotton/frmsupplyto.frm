VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsupplyto 
   Caption         =   "Form2"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14805
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   14805
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   840
      TabIndex        =   10
      Top             =   120
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
         TabIndex        =   12
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
         Left            =   5160
         TabIndex        =   11
         Top             =   360
         Width           =   9975
      End
      Begin VB.Image Image2 
         Height          =   9420
         Left            =   -720
         Picture         =   "frmsupplyto.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   20880
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D1FAFA&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   6360
      TabIndex        =   0
      Top             =   2880
      Width           =   8175
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
         Height          =   495
         Left            =   3840
         TabIndex        =   14
         Top             =   4440
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
         Height          =   495
         Left            =   5880
         TabIndex        =   13
         Top             =   4440
         Width           =   1695
      End
      Begin VB.ComboBox cmb_var 
         Height          =   315
         ItemData        =   "frmsupplyto.frx":1B6EBA
         Left            =   3000
         List            =   "frmsupplyto.frx":1B6EC7
         TabIndex        =   9
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txt_qty 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton cmdproceed 
         Caption         =   "PROCEED"
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
         Left            =   2040
         TabIndex        =   2
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton cmdabort 
         Caption         =   "ABORT"
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
         Left            =   360
         TabIndex        =   1
         Top             =   4440
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   375
         Left            =   6000
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   65929217
         CurrentDate     =   41571
      End
      Begin VB.Label lblvar 
         BackStyle       =   0  'Transparent
         Caption         =   "VARIETY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblqty 
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lbldate 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblsuply 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIED TO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Image Image1 
      Height          =   30000
      Left            =   0
      Picture         =   "frmsupplyto.frx":1B6ED7
      Top             =   0
      Width           =   45000
   End
End
Attribute VB_Name = "frmsupplyto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset
Dim flag As Integer
Dim z As Integer





Private Sub cmdback_Click()
Unload frmsupplyto
frmmain.Show
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdproceed_Click()
 Dim a As Integer
    Dim s As String
    Dim d As String
    Dim Y As String
    Dim g As Integer
    Dim f As Integer
    Dim qty As Integer
res.Open "select * from rawcotton where C_VARIETY='" & cmb_var.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    a = res(0)
    qty = res(2)
    res.Close
    f = 0
    If (qty < txt_qty.Text) Then
    f = 1
        MsgBox "CANNOT SUPPLY MORE THAN AVAILABLE "
     MsgBox "PLEASE FILL IN AGAIN"
    txt_qty.Text = ""
        
   ' g = Val(DTPicker1.Year)
   ' y=CStr(( g % 100))
    'g = Combopo.ListIndex
    'res.Open "select (S_QTY,C_id) from SUPPLIED_TO where p_id='" & Combopo.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
' d = res(0)
 'Y = res(g)
 '  d = DTPicker1.Day & "-" & s & "-" & DTPicker1.Year
'    Txtdate.Text = d
'  MsgBox ("DATE:" & Txtdate.Text)
'res.Open "select C_ID from rawcotton where C_VARIETY='" & cmb_var.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
   ' a = res(0)
   ' con.Execute ("update SUPPLIED_TO set S_day= '" & dtp1.Day & "',S_month='" & dtp1.Month & "',S_year='" & dtp1.Year & "',p_id='" & Combopo.Text & "' where S_QTY='" & d & "' and C_id='" & Y & "'")
       ' res.Close
        
   ' res.Open "select C_TOTQTY from rawcotton where C_VARIETY='" & cmb_var.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
's = res(0)
's = s - Val(txt_qty.Text)
'con.Execute ("UPDATE RAWCOTTON SET C_TOTQTY= '" & s & "' where C_ID='" & a & "'")
'Combopo.RemoveItem (g)
'Combopo.Refresh


'Combopo.Text = ""
'cmb_var.Text = ""
'txt_qty.Text = ""
Else

s = 0
res.Open "select P_ID from PRODUCTIONUNIT where P_TRASHQTY is NULL and P_SEEDQTY is null and P_LINTQTY is null", con, adOpenKeyset, adLockReadOnly, adCmdText
While (res.EOF = False)
s = s + 1
res.MoveNext
Wend
res.Close

If (s = 0) Then
   If (f = 0) Then
   
con.Execute ("insert into PRODUCTIONUNIT(STOCK_ID) values('1')")

res.Open "select P_ID from PRODUCTIONUNIT where P_TRASHQTY is NULL and P_SEEDQTY is null and P_LINTQTY is null", con, adOpenKeyset, adLockReadOnly, adCmdText
Y = res(0)
res.Close


      res.Open "select C_TOTQTY from rawcotton where C_ID='" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
s = res(0)
s = s - Val(txt_qty.Text)
con.Execute ("UPDATE RAWCOTTON SET C_TOTQTY= '" & s & "' where C_ID='" & a & "'")
 res.Close
 
    con.Execute ("insert into SUPPLIED_TO(S_day ,S_month,S_year,S_QTY,C_ID,P_ID) values('" & dtp1.Day & "','" & dtp1.Month & "','" & dtp1.Year & "','" & txt_qty.Text & "','" & a & "','" & Y & "')")
        
'res.Open "select P_ID from SUPPLIED_TO where S_day='" & dtp1.Day & "'s And S_month='" & dtp1.Month & "' And S_year='" & dtp1.Year & "' And S_QTY='" & txt_qty.Text & "' And C_ID='" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'd = res(0)
'res.Close
 MsgBox ("Quantity Successfully Suppied to Processing Unit  ")
    Else
    MsgBox "YOUR EARLIER QUANTITY SUPPLIED IS NOT PROCESSED TILL DATE>>>>>>>>>>>>>>>>>"
    End If
   End If
   End If

End Sub



Private Sub Form_Load()
    
    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    
    con.Open "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True;datasource=ADITYA"
    MsgBox ("Connection Established..........")
    flag = 1
End Sub

Private Sub txt_qty_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub
