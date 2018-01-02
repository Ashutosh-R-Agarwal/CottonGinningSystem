VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmpurchase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order"
   ClientHeight    =   7395
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txttotal 
      Enabled         =   0   'False
      Height          =   645
      Left            =   8760
      TabIndex        =   27
      ToolTipText     =   "DO NOT FILL"
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "BACK"
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
      Left            =   3000
      TabIndex        =   26
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
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
      Left            =   5520
      TabIndex        =   25
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdadd 
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
      Height          =   735
      Left            =   720
      TabIndex        =   23
      Top             =   6360
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   -1920
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "PurchaseOrder"
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
   Begin VB.Frame Purchase 
      BackColor       =   &H00CFFCFC&
      Caption         =   "PURCHASE ORDER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4815
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox txtagid 
         Height          =   315
         Left            =   2880
         TabIndex        =   22
         Top             =   480
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C4E9F9&
         Caption         =   "Y1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   7920
         TabIndex        =   12
         Top             =   1680
         Width           =   1815
         Begin VB.TextBox txtrpq3 
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txty1qty 
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txty1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "DO NOT FILL"
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkY1 
            BackColor       =   &H00CFEDFC&
            Caption         =   "Y1"
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C4E9F9&
         Caption         =   "S4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   5400
         TabIndex        =   11
         Top             =   1680
         Width           =   1695
         Begin VB.TextBox txtrpq2 
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txts4qty 
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txts4 
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "DO NOT FILL"
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CheckBox chkS4 
            BackColor       =   &H00CFEDFC&
            Caption         =   "S4"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C4E9F9&
         Caption         =   "BT "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   2880
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
         Begin VB.TextBox txtrpq1 
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtbtqty 
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtbt 
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   16
            ToolTipText     =   "DO NOT FILL"
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CheckBox chkbt 
            BackColor       =   &H00CFEDFC&
            Caption         =   "BT"
            Height          =   375
            Left            =   360
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7920
         TabIndex        =   1
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   65929217
         CurrentDate     =   41539
      End
      Begin VB.Label lblrpq 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RATE PER QUINTAL"
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
         Left            =   120
         TabIndex        =   29
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label lblEmname 
         BackStyle       =   0  'Transparent
         Caption         =   "PURCHASE REPRESENTATIVE"
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
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lbldate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
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
         Left            =   6240
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblnm 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblvar 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblrate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RATE"
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
         Left            =   -120
         TabIndex        =   3
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label lblqty 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.Label lbltotal 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6600
      TabIndex        =   28
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   30000
      Left            =   -600
      Picture         =   "frmpurchase.frx":0000
      Top             =   -600
      Width           =   45000
   End
   Begin VB.Label lblttl1 
      BackColor       =   &H00E4D2BA&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9840
      TabIndex        =   8
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblttl 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "frmpurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res1 As ADODB.Recordset
Dim con As ADODB.Connection
Dim res As ADODB.Recordset
Dim flag As Integer

Private Sub cmbvar_Click()
    If cmbvar.Text = "BT" Then
    txtrt.Text = 1200
    ElseIf cmbvar.Text = "S4" Then
    txtrt.Text = 1000
    ElseIf cmbvar.Text = "Y1" Then
    txtrt.Text = 800
    End If

End Sub


Private Sub chkbt_Click()

If (chkbt.Value = 1) Then
    txtbtqty.Visible = True
    txtbt.Visible = True
     txtrpq1.Visible = True
End If
If (chkbt.Value = 0) Then
    txtbtqty.Visible = False
    txtbt.Visible = False
     txtrpq1.Visible = False
End If


End Sub

Private Sub chkS4_Click()
If (chkS4.Value = 1) Then
    txts4qty.Visible = True
    txts4.Visible = True
     txtrpq2.Visible = True
    
End If
If (chkS4.Value = 0) Then
    txts4qty.Visible = False
    txts4.Visible = False
     txtrpq2.Visible = False
End If

End Sub

Private Sub chkY1_Click()
If (chkY1.Value = 1) Then
    txty1qty.Visible = True
    txty1.Visible = True
    txtrpq3.Visible = True
End If
If (chkY1.Value = 0) Then
    txty1qty.Visible = False
    txty1.Visible = False
     txtrpq3.Visible = False
End If

End Sub

Private Sub cmdadd_Click()
If txtagid.Text = "" Then
    MsgBox ("please Fill the Details...")
'If chkbt.Value = 1 Then
 '   If txtbt.Text = "" Or txtbtqty.Text = "" Then
  '  MsgBox ("PLEASE FILL THE DETAILS")
   ' End If
'If' chkS4.Value = 1 Then
    'If txts4.Text = "" Or txts4qty.Text = "" Then
    'MsgBox ("PLEASE FILL THE DETAILS")
    'End If
'If chkY1.Value = 1 Then
'    If txty1.Text = "" Or txty1qty.Text = "" Then
 '   MsgBox ("PLEASE FILL THE DETAILS")
  '  End If

Else

    
    Dim a As Integer
    Dim b As Integer
    Dim sum As Integer
    Dim rate As Integer
    Dim qty As Integer
    
    sum = 0
res1.Open "SELECT E_id from EMPLOYEE where E_name='" & Text1.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
   
    a = res1(0)
    res1.Close
    
    con.Execute ("insert into PURCHASEORDER (p_day,p_month,p_year,E_ID) values('" & DTPicker1.Day & "','" & DTPicker1.Month & "','" & DTPicker1.Year & "','" & a & "')")
   
    If (chkbt.Value = 1) Then
MsgBox ("IN BT")
rate = Val(txtrpq1.Text)
qty = Val(txtbtqty.Text)
txtbt.Text = (rate * qty)

    sum = sum + Val(txtbt.Text)
  ' MsgBox ("PURCHASE ODER ID GENDERATED IS:" & a)
   res1.Open "SELECT pono from PURCHASEORDER where p_day='" & DTPicker1.Day & "'and p_month='" & DTPicker1.Month & "' and p_year='" & DTPicker1.Year & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    a = res1(0)
    
   con.Execute ("insert into PURCHASE_QTY(PO_NO,P_QTY,P_RATE,P_VARIETY,RATE_QUINTAL) values('" & a & "','" & txtbtqty.Text & "','" & txtbt.Text & "','" & "BT" & "','" & rate & "')")
   res1.Close
   'MsgBox ("RECORD INSERTED SUCCESSFULLY")
    End If
    
If (chkS4.Value = 1) Then
    MsgBox ("IN S4")
    res1.Open "SELECT pono from PURCHASEORDER where p_day='" & DTPicker1.Day & "'and p_month='" & DTPicker1.Month & "'and  p_year='" & DTPicker1.Year & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    a = res1(0)
   rate = Val(txtrpq2.Text)
qty = Val(txts4qty.Text)
txts4.Text = (rate * qty)


   res1.Close
   con.Execute ("insert into PURCHASE_QTY(PO_NO,P_QTY,P_RATE,P_VARIETY,RATE_QUINTAL) values('" & a & "','" & txts4qty.Text & "','" & txts4.Text & "','" & "S4" & "','" & txtrpq2.Text & "')")
     sum = sum + Val(txts4.Text)
    End If
If (chkY1.Value = 1) Then
    MsgBox ("IN y1")
    
    
    rate = Val(txtrpq3.Text)
qty = Val(txty1qty.Text)
txty1.Text = (rate * qty)

    res1.Open "SELECT pono from PURCHASEORDER where p_day='" & DTPicker1.Day & "'and p_month='" & DTPicker1.Month & "' and p_year='" & DTPicker1.Year & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    a = res1(0)
    res1.Close
   con.Execute ("insert into PURCHASE_QTY(PO_NO,P_QTY,P_RATE,P_VARIETY,RATE_QUINTAL) values('" & a & "','" & txty1qty.Text & "','" & txty1.Text & "','" & "Y1" & "','" & txtrpq3.Text & "')")
   sum = sum + Val(txty1.Text)
    End If
  
   MsgBox ("PURCHASE ODER ID GENDERATED IS:" & a)
   res.Open "SELECT A_id from Agent where A_name='" & txtagid.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    b = Val(res(0))
   res.Close
   con.Execute ("insert into DEPENDS (A_ID,PO_NO) values('" & b & "','" & a & "')")
  
   txtbtqty.Visible = False
    txtbt.Visible = False

    txts4qty.Visible = False
    txts4.Visible = False
 
    txty1qty.Visible = False
    txty1.Visible = False
    txtagid.Text = ""
    Text1.Text = ""
    txtrpq3.Visible = False
      txtrpq2.Visible = False
       txtrpq1.Visible = False
    If (sum <> 0) Then
    txttotal.Text = sum
con.Execute ("update  PURCHASEORDER set P_TOTAMT='" & sum & "' where PONO='" & a & "'")
  Else
  con.Execute ("delete from   PURCHASEORDER  where PONO='" & a & "'")
  MsgBox ("YOU WILL HAVE TO SELECT SOMETHING TO CREATE A PURCHASE ODER")
  
  End If
  End If
End Sub

Private Sub txtqty_Change()
lblttl1.Caption = CInt(txtqty.Text) * CInt(txtrt.Text)
End Sub

Private Sub cmdback_Click()
Unload frmpurchase
frmmain.Show
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    Set res1 = New ADODB.Recordset
    con.Open ("Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True")
    MsgBox ("Conection Established ")
    flag = 1
    chkS4.Value = 0
    chkY1 = 0
    txtbtqty.Visible = False
    txtbt.Visible = False

    txts4qty.Visible = False
    txts4.Visible = False
 
    txty1qty.Visible = False
    txty1.Visible = False
     txtrpq3.Visible = False
      txtrpq2.Visible = False
       txtrpq1.Visible = False
End Sub

Private Sub txtagid_Click()
res1.Open "SELECT E_NAME from EMPLOYEE where E_ID=(select E_ID from OFFICEWORKER where E_JOBTITLE = 'PURCHASE REPRESENTATIVE')", con, adOpenKeyset, adLockReadOnly, adCmdText
   
Text1.Text = res1(0)
res1.Close

End Sub

Private Sub txtagid_DropDown()
If (flag = 1) Then

 flag = 2
 res.Open "select distinct(A_NAME) from AGENT", con, adOpenKeyset, adLockReadOnly, adCmdText
    While res.EOF <> True
         txtagid.AddItem (res(0))
    res.MoveNext
    Wend
res.Close
End If
End Sub

Private Sub txtbt_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtbtqty_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub



Private Sub txts4_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txts4qty_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub


Private Sub txty1_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txty1qty_KeyPress(KeyAscii As Integer)
Dim key As String
key = Chr$(KeyAscii) 'ascii no to character string
If ((key < "0" Or key > "9") And key <> ":") Then
KeyAscii = 0
Beep
End If

End Sub
