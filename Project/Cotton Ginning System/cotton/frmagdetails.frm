VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmagdetails 
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   360
      Top             =   5040
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
   Begin VB.ComboBox cmbsearch 
      Height          =   315
      ItemData        =   "frmagdetails.frx":0000
      Left            =   5280
      List            =   "frmagdetails.frx":000A
      TabIndex        =   2
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E4BAC4&
      Caption         =   "AGENT"
      Height          =   6015
      Left            =   3000
      TabIndex        =   0
      Top             =   2880
      Width           =   13095
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmagdetails.frx":0024
         Height          =   2295
         Left            =   360
         TabIndex        =   1
         Top             =   2040
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblsearch 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
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
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   11640
      Left            =   0
      Picture         =   "frmagdetails.frx":0039
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20280
   End
End
Attribute VB_Name = "frmagdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim str As String

Private Sub cmdmodify_Click()
DataGrid1.AllowUpdate = True
End Sub

Private Sub cmbsearch_Click()
Dim a As Integer
Dim cnt As Integer
Dim i As Integer
cnt = 1
str = cmbsearch.Text

If str = "Agent_ID" Then
    a = InputBox("Enter The Agent ID : ")
    res.Open "SELECT A_id from Agent", con, adOpenKeyset, adLockReadOnly, adCmdText
   While res.EOF = False
    i = i + 1
    res.MoveNext
    Wend
    res.Close
   res.Open "SELECT A_id from Agent", con, adOpenKeyset, adLockReadOnly, adCmdText
    While cnt <> i And res(0) <> a
    cnt = cnt + 1
    res.MoveNext
    Wend
    If (i = cnt) Then
   MsgBox ("RECORD NOT FOUND")
   Else
    DataGrid1.Refresh
    
    DataGrid1.AllowArrows = True
    DataGrid1.TabIndex = cnt
    MsgBox ("RECORD AT NUMBER  " & cnt)
    End If
Else
    a = InputBox("Enter The Contact Number :")
    

    res.Open "SELECT A_Phno from Agent", con, adOpenKeyset, adLockReadOnly, adCmdText
   While res.EOF = False
    i = i + 1
    res.MoveNext
    Wend
    res.Close
   res.Open "SELECT A_Phno from Agent", con, adOpenKeyset, adLockReadOnly, adCmdText
    While cnt <> i And res(0) <> a
    cnt = cnt + 1
    res.MoveNext
    Wend
    If (i = cnt) Then
   MsgBox ("RECORD NOT FOUND")
   Else
    DataGrid1.Refresh
    
    DataGrid1.AllowArrows = True
    DataGrid1.TabIndex = cnt
    MsgBox ("RECORD AT NUMBER  " & cnt)
    End If
End If
res.Close

End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    
    con.Open "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True"
    MsgBox ("Connection Established..........")
    
    'con.Execute ("Select * from Agent")
End Sub
