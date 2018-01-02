VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frminvoice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INVOICE"
   ClientHeight    =   5250
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
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
      Left            =   9240
      TabIndex        =   25
      Top             =   4680
      Width           =   1935
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
      Height          =   495
      Left            =   3120
      TabIndex        =   22
      Top             =   4680
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
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   4680
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6240
      Top             =   5400
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "invoice"
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
   Begin VB.CommandButton cmdadd 
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
      Left            =   840
      TabIndex        =   9
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Frame INVOICE 
      BackColor       =   &H00CFFCFC&
      Caption         =   "INVOICE"
      Height          =   4215
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   11535
      Begin VB.TextBox cmbvar 
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txttot 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2640
         TabIndex        =   24
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txtrt 
         Height          =   375
         Left            =   2640
         TabIndex        =   20
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox lblcid1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   495
         Left            =   8160
         TabIndex        =   19
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox lblcadd1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   855
         Left            =   8160
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox lblccn1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   495
         Left            =   8160
         TabIndex        =   17
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox lblcp1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   495
         Left            =   8160
         TabIndex        =   16
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtqty 
         Height          =   405
         Left            =   2640
         TabIndex        =   2
         Top             =   2880
         Width           =   2775
      End
      Begin VB.ComboBox cmbnm 
         Height          =   315
         ItemData        =   "frminvoice.frx":0000
         Left            =   2640
         List            =   "frminvoice.frx":0002
         TabIndex        =   1
         Top             =   1080
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20512769
         CurrentDate     =   41539
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY ID."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   15
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lbladd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY ADDRESS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         TabIndex        =   14
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lblmn 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblcp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT PERSON"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   12
         Top             =   360
         Width           =   2295
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
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblrate 
         BackStyle       =   0  'Transparent
         Caption         =   "RATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
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
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblnm 
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY NAME"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lbldate 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   30000
      Left            =   -360
      Picture         =   "frminvoice.frx":0004
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   45000
   End
   Begin VB.Label lblttl1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblttl 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   6600
      Width           =   1455
   End
End
Attribute VB_Name = "frminvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset
Dim flag1 As Integer
Dim flag As Integer
Dim str As String
Dim comid As Integer
Option Explicit

Private Sub cmbnm_DropDown()
If (flag1 = 1) Then
 res.Open "select distinct(COM_NAME) from COMPANY ", con, adOpenKeyset, adLockReadOnly, adCmdText
    While res.EOF <> True
         cmbnm.AddItem (res(0))
    res.MoveNext
    Wend
res.Close
flag1 = 2
End If
End Sub

Private Sub cmbpoid_DropDown()
If (flag = 1) Then
 res.Open "select distinct(P_ID) from PRODUCTIONUNIT ", con, adOpenKeyset, adLockReadOnly, adCmdText
    While res.EOF <> True
         cmbpoid.AddItem (res(0))
    res.MoveNext
    Wend
res.Close
flag = 2
End If
End Sub

Private Sub cmbnm_Click()
    str = cmbnm.Text
    Dim a As Integer
    Dim var As String
    MsgBox "YOU HAVE SELECTED COMPANY:" & str
    
   res.Open "SELECT * from COMPANY where COM_NAME='" & str & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
   a = res(0)
   var = res(4)
   If (var = "LINT COMPANY") Then
   
   cmbvar.Text = "LINT"
   End If
   If (var = "OIL COMPANY") Then
   cmbvar.Text = "BEENS"
   End If
   str = res(3)
    lblcadd1.Text = res(3)
    lblccn1.Text = res(2)
   comid = a
    res1.Open "SELECT * from CONTACTPERSON where COM_ID='" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
   
    lblcid1.Text = res(0)
    
    lblcp1.Text = res1(0)
    res.Close
    res1.Close
    
    
End Sub



Private Sub cmbvar_Click()
If cmbvar.Text = "BEENS" Then
txtrt.Text = 1200
ElseIf cmbvar.Text = "LINT" Then
txtrt.Text = 1000
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


Private Sub txtnmp_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case vbKey0 To vbKey9
KeyAscii = 0
Beep
End Select

End Sub

Private Sub cmdback_Click()
Unload frminvoice
frmmain.Show
End Sub

Private Sub cmdclear_Click()
cmbnm.Text = ""
cmbvar.Text = ""
txtrt.Text = ""
txtqty.Text = ""
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub Form_Load()
    
    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    
    con.Open "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True;datasource=ADITYA"
    MsgBox ("Connection Established..........")
    flag1 = 1
flag = 1
End Sub

Private Sub cmdadd_Click()
If cmbnm.Text = "" Or cmbvar.Text = "" Or txtrt.Text = "" Or txtqty.Text = "" Then
MsgBox ("PLEASE FILL THE DETAILS")
Else
 Dim a As Integer
    Dim s As String
    Dim d As String
    Dim Y As String
    Dim g As Integer
  Dim rate As Integer
  Dim qty As Integer
  Dim var As String
  Dim TOT As Integer
  Dim f As Integer
  
   ' g = Val(DTPicker1.Year)
   ' y=CStr(( g % 100))
   ' s = Month(DTPicker1.Month)
   res.Open "SELECT * from TOTALOUTPUTQTY", con, adOpenKeyset, adLockReadOnly, adCmdText
    d = res(0)
    Y = res(3)
    g = res(1)
    res.Close
    f = 0
    
    If (cmbvar.Text = "BEENS") Then
MsgBox "BEFORE QUANTITY  IN TOTALOUTPUT OF SELECTED ITEM" & d
    If (d < Val(txtqty.Text)) Then
    f = 1

    MsgBox "QUANTITY SPECIFIED IS MORE THAN THE AVAILABLE STOCK"
    MsgBox "PLEASE FILL THE STOCK AGAIN"
    txtqty.Text = ""
    End If
    
    End If
    If (cmbvar.Text = "LINT") Then
    If (Y < Val(txtqty.Text)) Then
    
    f = 1

    MsgBox "QUANTITY SPECIFIED IS MORE THAN THE AVAILABLE STOCK"
    MsgBox "PLEASE FILL THE STOCK AGAIN"
    txtqty.Text = ""

    End If
    End If
    If (cmbvar.Text = "TRASH") Then
    If (g < Val(txtqty.Text)) Then
    f = 1

    MsgBox "QUANTITY SPECIFIED IS MORE THAN THE AVAILABLE STOCK"
   MsgBox "PLEASE FILL THE STOCK AGAIN"
    txtqty.Text = ""
 
    End If
    
        End If
If (f = 0) Then

   res1.Open "SELECT COM_ID from COMPANY where COM_NAME='" & cmbnm.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
   res1.Close
  
  
    res1.Open "SELECT COM_ID from COMPANY where COM_NAME='" & cmbnm.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    a = Val(res1(0))
    
    
    
   ' res.Open "SELECT P_ID from PRODUCTIONUNIT where COM_NAME='" & cmbnm.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
 '  d = DTPicker1.Day & "-" & s & "-" & DTPicker1.Year
'    Txtdate.Text = d
'  MsgBox ("DATE:" & Txtdate.Text)


qty = Val(txtqty.Text)
rate = Val(txtrt.Text)
txttot.Text = (qty * rate)

    con.Execute ("insert  into INVOICE (I_DAY,I_MONTH,I_YEAR,I_QTYSUP,I_RATE,E_ID,COM_ID,VARIETY,STOCK_ID,I_TOTAL) values('" & DTPicker1.Day & "','" & DTPicker1.Month & "','" & DTPicker1.Year & "','" & txtqty.Text & "','" & txtrt.Text & "','1355','" & a & "','" & cmbvar.Text & "','1','" & txttot.Text & "')")
   res.Open "SELECT I_NO from INVOICE where I_DAY='" & DTPicker1.Day & "' and I_MONTH='" & DTPicker1.Month & "' and I_YEAR='" & DTPicker1.Year & "' and I_QTYSUP='" & txtqty.Text & "' and VARIETY='" & cmbvar.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    a = res(0)
    MsgBox ("INVOICE ID IS:" & a)
    res.Close
    If (cmbvar.Text = "BEENS") Then
  '  MsgBox "HII I M ASHUTOSH"

    d = d - Val(txtqty.Text)
    con.Execute ("update TOTALOUTPUTQTY set TOTAL_SEED='" & d & "'")
    'MsgBox "IN BEANS"
    End If
    If (cmbvar.Text = "LINT") Then
    
    Y = Y - Val(txtqty.Text)
    con.Execute ("update TOTALOUTPUTQTY set TOTAL_LINT='" & Y & "'")
    'MsgBox "IN LINT"
    End If
    If (cmbvar.Text = "TRASH") Then
    
    g = g - Val(txtqty.Text)
    con.Execute ("update TOTALOUTPUTQTY set TOTAL_TRASH='" & g & "'")
    'MsgBox "IN TRASH"
    End If
    MsgBox ("INSERTED SUCESSFULLY: >>>>>>>>>>")
    End If
    End If
    
End Sub


Private Sub txtqty_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtrt_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub
