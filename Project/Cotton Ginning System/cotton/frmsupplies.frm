VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmsupplies 
   Caption         =   "Form2"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   11400
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
      Height          =   495
      Left            =   11160
      TabIndex        =   22
      Top             =   9720
      Width           =   1455
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
      Left            =   12720
      TabIndex        =   21
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   1080
      TabIndex        =   18
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   360
         Width           =   9975
      End
      Begin VB.Image Image3 
         Height          =   9420
         Left            =   -960
         Picture         =   "frmsupplies.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   20880
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   -1440
      Top             =   4560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
      RecordSource    =   "PURCHASEORDER"
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
      BackColor       =   &H00D1FAFA&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7215
      Left            =   5880
      TabIndex        =   0
      Top             =   3600
      Width           =   8535
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
         Left            =   3600
         TabIndex        =   23
         Top             =   6120
         Width           =   1455
      End
      Begin VB.TextBox txt_agt 
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   2880
         Width           =   2775
      End
      Begin VB.ComboBox Combopo 
         Height          =   315
         Left            =   2640
         TabIndex        =   10
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtvar1 
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtvar2 
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtvar3 
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtqty1 
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtqty2 
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtqty3 
         Height          =   375
         Left            =   6240
         TabIndex        =   4
         Top             =   3840
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
         Left            =   1920
         TabIndex        =   2
         Top             =   6120
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
         Left            =   240
         TabIndex        =   1
         Top             =   6120
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   65994753
         CurrentDate     =   41571
      End
      Begin VB.Label lbl_agnm 
         BackStyle       =   0  'Transparent
         Caption         =   "AGENT NAME"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   17
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblpono 
         BackStyle       =   0  'Transparent
         Caption         =   "PURCHASE NO"
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
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblvar 
         BackStyle       =   0  'Transparent
         Caption         =   "VARIETY"
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
         Left            =   360
         TabIndex        =   14
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label lblqty 
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY"
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
         Left            =   240
         TabIndex        =   13
         Top             =   3840
         Width           =   1695
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
         Left            =   4920
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblsuply 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIES DETAILS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Image Image2 
      Height          =   30000
      Left            =   120
      Picture         =   "frmsupplies.frx":1B6EBA
      Top             =   0
      Width           =   45000
   End
   Begin VB.Image Image1 
      Height          =   30000
      Left            =   -720
      Top             =   -120
      Width           =   45000
   End
End
Attribute VB_Name = "frmsupplies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset
Dim a As Integer
Dim flag As Integer
Dim flag1 As Integer

Private Sub cmdabort_Click()
a = Combopo.ListIndex
con.Execute ("delete from PURCHASE_QTY where PO_NO ='" & Combopo.Text & "'")
con.Execute ("COMMIT")
con.Execute ("delete from depends where PO_NO ='" & Combopo.Text & "'")
con.Execute ("delete from PURCHASEORDER where PONO ='" & Combopo.Text & "'")
con.Execute ("COMMIT")
MsgBox ("TRANSACTION ABORTED SUCCESSFULLY..........")
            txtvar2.Text = ""
            txtqty2.Text = ""
            txtvar1.Text = ""
            txtqty1.Text = ""
            txtvar3.Text = ""
            txtqty3.Text = ""
            Combopo.Text = ""
            Combopo.Refresh
txtvar2.Text = ""
            txtqty2.Text = ""
            txtvar1.Text = ""
            txtqty1.Text = ""
            txtvar3.Text = ""
            txtqty3.Text = ""
            txt_agt.Text = ""

Combopo.Refresh
Combopo.RemoveItem (a)
End Sub

Private Sub cmdback_Click()
Unload frmsupplies
frmmain.Show
End Sub

Private Sub cmdclear_Click()
Combopo.Text = ""
txt_agt.Text = ""
txtqty1.Text = ""
txtqty2.Text = ""
txtqty.Text = ""
txtvar1.Text = ""
txtvar2.Text = ""
txtvar3.Text = ""
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdproceed_Click()
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim f As Integer
Dim t As Integer
Dim n As Integer
Dim day1 As Integer
Dim mn As Integer
Dim yr As Integer
Dim day2 As Integer
Dim mn2 As Integer
Dim yr2 As Integer

a = Combopo.ListIndex

res.Open "SELECT * from PURCHASEORDER where PONO='" & Combopo.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
day1 = res(1)
MsgBox "DAY!" & day1
mn = res(2)
MsgBox "MONTH!" & mn
yr = res(3)
MsgBox "YAER1" & yr
day2 = dtp1.Day
MsgBox "days 2" & day2
mn2 = dtp1.Month
MsgBox "month2 " & mn2
yr2 = dtp1.Year
MsgBox "year2" & yr2
res.Close
If (day1 >= day2 And mn >= mn2 And yr >= yr2) Then

res1.Open "SELECT A_id from Agent where A_name='" & txt_agt.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
f = res1(0)
res1.Close
If (txtqty1.Visible = True) Then
res.Open "SELECT * from RAWCOTTON where C_VARIETY='" & txtvar1.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
c = res(0)
b = res(2)
b = b + Val(txtqty1.Text)
con.Execute ("UPDATE RAWCOTTON SET C_TOTQTY='" & b & "' where C_ID='" & c & "'")
con.Execute ("UPDATE PURCHASEORDER SET PENDING='1' where PONO='" & Combopo.Text & "'")
con.Execute ("insert into SUPPLIES (C_ID,s_day,s_month,s_year,A_ID) values('" & res(0) & "','" & dtp1.Day & "','" & dtp1.Month & "','" & dtp1.Year & "','" & f & "')")
res.Close
MsgBox ("YES I HAVE DONE IT")
MsgBox "IN 1ST" & txtvar1.Text
MsgBox "QTY" & b

End If


If (txtqty2.Visible = True) Then
res.Open "SELECT * from RAWCOTTON where C_VARIETY='" & txtvar2.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
c = res(0)
t = res(2)
t = t + Val(txtqty2.Text)
con.Execute ("UPDATE RAWCOTTON SET C_TOTQTY='" & t & "' where C_ID='" & c & "'")
MsgBox ("IN TEXTBOX" & txtvar2.Text)

con.Execute ("UPDATE PURCHASEORDER SET PENDING='1' where PONO='" & Combopo.Text & "'")
con.Execute ("insert into SUPPLIES (C_ID,s_day,s_month,s_year,A_ID) values('" & res(0) & "','" & dtp1.Day & "','" & dtp1.Month & "','" & dtp1.Year & "','" & f & "')")
res.Close

MsgBox ("YES I HAVE DONE IT")
MsgBox "IN 1ST" & txtvar2.Text
MsgBox "QTY" & t

End If
If (txtqty3.Visible = True) Then
res.Open "SELECT * from RAWCOTTON where C_VARIETY='" & txtvar3.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
c = res(0)
n = res(2)
n = n + Val(txtqty3.Text)
con.Execute ("UPDATE RAWCOTTON SET C_TOTQTY='" & n & "' where C_ID='" & c & "'")
con.Execute ("UPDATE PURCHASEORDER SET PENDING='1' where PONO='" & Combopo.Text & "'")
'res1.Open "SELECT * from PURCHASEORDER where PONO='" & Combopo.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
con.Execute ("insert into SUPPLIES (C_ID,s_day,s_month,s_year,A_ID) values('" & res(0) & "','" & dtp1.Day & "','" & dtp1.Month & "','" & dtp1.Year & "','" & f & "')")
res.Close
MsgBox ("YES I HAVE DONE IT")
MsgBox "IN 1ST" & txtvar3.Text
MsgBox "QTY" & n

End If
txtvar2.Text = ""
            txtqty2.Text = ""
            txtvar1.Text = ""
            txtqty1.Text = ""
            txtvar3.Text = ""
            txtqty3.Text = ""
            txt_agt.Text = ""
Combopo.Refresh
Combopo.Refresh

    
Combopo.RemoveItem (a)
Else
MsgBox ("DATE OF SUPPLIES CANNOT BE GREATER THAN PURCHASE ORDER..........")
txtvar2.Text = ""
            txtqty2.Text = ""
            txtvar1.Text = ""
            txtqty1.Text = ""
            txtvar3.Text = ""
            txtqty3.Text = ""
            Combopo.Text = ""
            Combopo.Refresh
txtvar2.Text = ""
            txtqty2.Text = ""
            txtvar1.Text = ""
            txtqty1.Text = ""
            txtvar3.Text = ""
            txtqty3.Text = ""
            txt_agt.Text = ""

End If
End Sub

Private Sub Combopo_Click()
    Dim b As Integer
    a = Combopo.Text
    MsgBox ("YOu have creted " & a)
    res.Open "SELECT * from DEPENDS where PO_NO='" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    ' If (res.EOF = False) Then
     If (res.EOF = False) Then
     b = res(0)
    res.Close
    res.Open "SELECT A_NAME from AGENT where A_ID='" & b & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    txt_agt.Text = res(0)
    res.Close
    
    res.Open "SELECT * from PURCHASE_QTY where PO_NO='" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    flag = 0
    Combopo.Refresh
    
    MsgBox ("INWERGRGREG")
    txtvar2.Text = ""
            txtqty2.Text = ""
            txtvar1.Text = ""
            txtqty1.Text = ""
            txtvar3.Text = ""
            txtqty3.Text = ""
            
    While res.EOF <> True
        If (flag = 0 And res.EOF = False) Then
            txtvar1.Text = res(3)
            txtqty1.Text = res(1)
            flag = 1
            res.MoveNext
        End If
        If (flag = 1 And res.EOF = False) Then
        txtvar2.Text = res(3)
            txtqty2.Text = res(1)
            flag = 2
            res.MoveNext
        End If
        If (flag = 2 And res.EOF = False) Then
        txtvar3.Text = res(3)
            txtqty3.Text = res(1)
            flag = 3
            res.MoveNext
        End If
        
    Wend
    
res.Close
If (flag = 1) Then
txtvar2.Visible = False
            txtqty2.Visible = False
            txtvar3.Visible = False
            txtqty3.Visible = False
End If

If (flag = 2) Then
            txtvar3.Visible = False
            txtqty3.Visible = False
End If
Else
MsgBox "CURREENT RECORD HAS BEEN DELETED"
res.Close

End If

'Else
'MsgBox ("SOMTHNG WENT WRONG")
'End If


End Sub

Private Sub Combopo_DropDown()
If (flag1 = 1) Then
 res.Open "select distinct(PONO) from purchaseoRder where PENDING= '0' ", con, adOpenKeyset, adLockReadOnly, adCmdText
    While res.EOF <> True
         Combopo.AddItem (res(0))
    res.MoveNext
    Wend
res.Close
flag1 = 2
End If



End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    Set res1 = New ADODB.Recordset
    con.Open "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True"
    MsgBox ("Connection Established..........")
flag1 = 1
End Sub

Private Sub txtqty1_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtqty2_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtqty3_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub
