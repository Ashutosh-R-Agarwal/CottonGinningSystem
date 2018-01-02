VERSION 5.00
Begin VB.Form frmprocess 
   Caption         =   "Form1"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19095
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10470
   ScaleWidth      =   19095
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   960
      TabIndex        =   17
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   360
         Width           =   9975
      End
      Begin VB.Image Image2 
         Height          =   9420
         Left            =   -600
         Picture         =   "frmprocess.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   20880
      End
   End
   Begin VB.Timer tmr1 
      Interval        =   500
      Left            =   960
      Top             =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7215
      Left            =   5040
      TabIndex        =   0
      Top             =   3000
      Width           =   10095
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
         Left            =   5400
         TabIndex        =   21
         Top             =   6480
         Width           =   1815
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
         Left            =   7800
         TabIndex        =   20
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00A3F1EF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5295
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   7935
         Begin VB.TextBox Cmbvar 
            Enabled         =   0   'False
            Height          =   375
            Left            =   3720
            TabIndex        =   16
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox Combopo 
            Enabled         =   0   'False
            Height          =   495
            Left            =   3720
            TabIndex        =   15
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtstkname 
            Enabled         =   0   'False
            Height          =   375
            Left            =   3720
            TabIndex        =   8
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox txtb 
            Height          =   375
            Left            =   3720
            TabIndex        =   7
            Top             =   2760
            Width           =   3375
         End
         Begin VB.TextBox txtl 
            Height          =   375
            Left            =   3720
            TabIndex        =   6
            Top             =   3600
            Width           =   3375
         End
         Begin VB.TextBox txttr 
            Height          =   375
            Left            =   3720
            TabIndex        =   5
            Top             =   4440
            Width           =   3375
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
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblstkid 
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCTION ID"
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
            TabIndex        =   13
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label lblstknm 
            BackStyle       =   0  'Transparent
            Caption         =   "COTTON INPUT"
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
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblqty 
            BackStyle       =   0  'Transparent
            Caption         =   "QUANTITY OF BEENS"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   11
            Top             =   2760
            Width           =   2415
         End
         Begin VB.Label lbll 
            BackStyle       =   0  'Transparent
            Caption         =   "QUANTITY OF LINT"
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
            Left            =   120
            TabIndex        =   10
            Top             =   3600
            Width           =   2895
         End
         Begin VB.Label lbltr 
            BackStyle       =   0  'Transparent
            Caption         =   "QUANTITY OF TRASH"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   9
            Top             =   4440
            Width           =   2895
         End
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   3
         Top             =   6480
         Width           =   1575
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   6480
         Width           =   1455
      End
   End
   Begin VB.Label lblprocess 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESS DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   6360
      TabIndex        =   1
      Top             =   2160
      Width           =   8055
   End
   Begin VB.Image Image1 
      Height          =   10935
      Left            =   120
      Picture         =   "frmprocess.frx":1B6EBA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "frmprocess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset
Dim res2 As New ADODB.Recordset
Dim flag As Integer
Dim g As Integer

Private Sub cmdsubmit_Click()
If txtid.Text = "" Or txtemp.Text = "" Or txthrs.Text = "" Or txtvar.Text = "" Or txtstkname.Text = "" Or txtqty.Text = "" Then
MsgBox "PLEASE FILL DETAILS"
End If
frmemployee.Show
End Sub



Private Sub cmdadd_Click()
   Dim a As Integer
    Dim s As String
    Dim d As String
    Dim Y As String

  Dim h As Integer
  Dim i As Integer
  Dim j As Integer
  res.Open "select P_ID from PRODUCTIONUNIT where P_TRASHQTY is NULL and P_SEEDQTY is null and P_LINTQTY is null", con, adOpenKeyset, adLockReadOnly, adCmdText
   If res.EOF <> True Then
  res.Close
  
  res.Open "select P_ID from PRODUCTIONUNIT where P_TRASHQTY is NULL and P_SEEDQTY is null and P_LINTQTY is null", con, adOpenKeyset, adLockReadOnly, adCmdText
g = res(0)
MsgBox ("YOU HAVE SELECTED:" & g)


  con.Execute ("UPDATE  PRODUCTIONUNIT SET P_TRASHQTY='" & txttr.Text & "',P_SEEDQTY='" & txtb.Text & "',P_LINTQTY='" & txtl.Text & "'")

res.Close

'res.Open "select * from rawcotton where C_VARIETY='" & cmbvar.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
  '  a = res(0)
    'd = res(2)
    'd = d - Val(txtstkname.Text)
    'con.Execute ("UPDATE  rawcotton SET C_TOTQTY= '" & d & "'where C_VARIETY='" & Cmbvar.Text & "'")
    
   ' con.Execute ("insert into SUPPLIED_TO (C_ID,S_QTY) values('" & a & "','" & txtstkname.Text & "')")
    'con.Execute ("insert into PRODUCTIONUNIT (P_ID,P_TRASHQTY,P_SEEDQTY,P_LINTQTY,STOCK_ID) values('" & txtid.Text & "','" & txttr.Text & "','" & txtb.Text & "','" & txtl.Text & "',1)")
   're 's.Close
    're
    's.Open "SELECT P_ID from PRODUCTIONUNIT where P_TRASHQTY='" & txttr.Text & "' and P_SEEDQTY='" & txtb.Text & "' and P_LINTQTY='" & txtl.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'    res1.Open "SELECT A_id from Agent where A_name='" & txtnm.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'    a = Val(res1(0))
    res1.Open "SELECT * from TOTALOUTPUTQTY ", con, adOpenKeyset, adLockReadOnly, adCmdText
    h = Val(res1(0))
    MsgBox ("TOTAL_SEED TILL DATE IN TOTALOUTPUT:" & h)
    i = Val(res1(1))
    MsgBox ("TOTAL_TRASH TILL DATE IN TOTALOUTPUT:" & i)
    j = Val((res1(3)))
    MsgBox ("TOTAL_LINT TILL DATE IN TOTALOUTPUT:" & j)
    h = h + Val(txtb.Text)
    i = i + Val(txttr.Text)
    
    j = j + Val(txtl.Text)
    con.Execute ("UPDATE  TOTALOUTPUTQTY SET TOTAL_SEED= '" & h & "', TOTAL_TRASH= '" & i & "',TOTAL_LINT = '" & j & "'")
   ' MsgBox ("YOUR AGENT NAME: " & a)
   
   
   MsgBox (" Successfully PROCESSED ......")
   Else
   res.Close
    MsgBox "NO QUANTITY TO PROCESS PLEASE GO BACK TO SUPPLIES TO "
    frmsupplyto.Show
    End If
    
End Sub

Private Sub Combopo_DropDown()
If (flag = 1) Then

 res.Open "select P_ID from PRODUCTIONUNIT where P_TRASHQTY is NULL and P_SEEDQTY is null and P_LINTQTY is null", con, adOpenKeyset, adLockReadOnly, adCmdText
    While res.EOF <> True
         Combopo.AddItem (res(0))
    res.MoveNext
    Wend
res.Close
flag = 2
End If


End Sub


Private Sub cmdback_Click()
Unload frmprocess
frmmain.Show
frmmain.Show
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    Set res1 = New ADODB.Recordset
     Set res2 = New ADODB.Recordset
    con.Open ("Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True;datasource=ADITYA")
    MsgBox ("Conection Established ")

    res.Open "select P_ID from PRODUCTIONUNIT where P_TRASHQTY is NULL and P_SEEDQTY is null and P_LINTQTY is null", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    If res.EOF <> True Then
    
    g = res(0)
    Combopo.Text = g
         res1.Open "select C_ID,S_QTY from SUPPLIED_TO where P_ID='" & g & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
txtstkname.Text = res1(1)
flag = res1(0)
res2.Open "select * from rawcotton where C_ID='" & flag & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
Cmbvar.Text = res2(1)
res.Close
res2.Close
res1.Close

Else
MsgBox "THEIR IS  NO QUANTY TO PROCESS PLEASE GO BACK TO SUPPLY TO"
res.Close
End If



End Sub

Private Sub tmr1_Timer()
If lblprocess.ForeColor = &H400040 Then
lblprocess.ForeColor = &HFFFF&
ElseIf lblprocess.ForeColor = &HFFFF& Then
lblprocess.ForeColor = &H400040
End If
End Sub


