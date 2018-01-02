VERSION 5.00
Begin VB.Form frmemployee 
   Caption         =   "Form1"
   ClientHeight    =   9960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   13410
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
      Left            =   17280
      TabIndex        =   29
      Top             =   9000
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
      Left            =   14400
      TabIndex        =   28
      Top             =   9000
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   960
      TabIndex        =   25
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   360
         Width           =   9975
      End
      Begin VB.Image Image3 
         Height          =   9420
         Left            =   -600
         Picture         =   "frmemployee.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   20880
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7215
      Left            =   14280
      TabIndex        =   19
      Top             =   3720
      Width           =   6735
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00808080&
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
         Height          =   1095
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdclear 
         BackColor       =   &H00808080&
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
         Height          =   1095
         Left            =   360
         MaskColor       =   &H00808080&
         TabIndex        =   23
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdupdate 
         BackColor       =   &H00808080&
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1680
         TabIndex        =   22
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H00808080&
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3000
         TabIndex        =   21
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00808080&
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3000
         TabIndex        =   20
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   30000
         Left            =   0
         Picture         =   "frmemployee.frx":1B6EBA
         Top             =   0
         Width           =   45000
      End
   End
   Begin VB.Timer Timer1 
      Left            =   14040
      Top             =   7920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDCEE0&
      BorderStyle     =   0  'None
      Caption         =   "EMPLOYEE ID"
      ForeColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   1680
      TabIndex        =   0
      Top             =   3720
      Width           =   12255
      Begin VB.ComboBox combojob 
         Height          =   315
         ItemData        =   "frmemployee.frx":310C25
         Left            =   8520
         List            =   "frmemployee.frx":310C32
         TabIndex        =   17
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox txtsal 
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   5280
         Width           =   2175
      End
      Begin VB.ComboBox ComboGender 
         Height          =   315
         ItemData        =   "frmemployee.frx":310C6E
         Left            =   2520
         List            =   "frmemployee.frx":310C78
         TabIndex        =   13
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtdob 
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         ToolTipText     =   "04-JAN-1997"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox Combodes 
         Height          =   315
         ItemData        =   "frmemployee.frx":310C82
         Left            =   8400
         List            =   "frmemployee.frx":310C8C
         TabIndex        =   11
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox txtemail 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   9
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox txtaddr 
         Appearance      =   0  'Flat
         Height          =   1335
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox txtnm 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblperday 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   9120
         TabIndex        =   18
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label lblpost 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "JOB TITLE"
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
         Left            =   6360
         TabIndex        =   16
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label lblsal 
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY"
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
         TabIndex        =   14
         Top             =   5280
         Width           =   1455
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
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lbldesgn 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DESIGNATION"
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
         Left            =   6240
         TabIndex        =   5
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lblemail 
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
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
         Left            =   600
         TabIndex        =   4
         Top             =   4440
         Width           =   1455
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
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   2760
         Width           =   1815
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
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   1920
         Width           =   1935
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
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label lblemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE DETAILS"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8160
      TabIndex        =   10
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   30000
      Left            =   -960
      Picture         =   "frmemployee.frx":310CAE
      Top             =   -120
      Width           =   45000
   End
End
Attribute VB_Name = "frmemployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset

Dim str As String



Private Sub cmdadd_Click()
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    
    If txtnm.Text = "" Or txtaddr.Text = "" Or txtemail.Text = "" Or ComboGender.Text = " " Then
        MsgBox "PLEASE FILL THE DETAILS"
    End If
    
    str = Combodes.Text
    
    con.Execute ("insert into Employee(E_name,E_dob,E_gender,E_add,E_email,E_designation) values('" & txtnm.Text & "','" & txtdob.Text & "','" & ComboGender.Text & "','" & txtaddr.Text & "','" & txtemail.Text & "','" & Combodes.Text & "')")
    
    con.Execute ("commit")
    
    res1.Open "SELECT E_id from Employee where E_name='" & txtnm.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    a = Val(res1(0))
    
    If str = "Office Worker" Then
        con.Execute ("insert into OfficeWorker(E_id,E_jobtitle,E_sal) values('" & a & "','" & combojob.Text & "','" & txtsal.Text & "')")
    Else
        b = 100 * CInt(txtnohrs.Text)
        lblperday.Caption = b
        con.Execute ("insert into OnsiteWorker(E_id,E_nohrs,E_shift,E_wages) values('" & a & "','" & txtnohrs.Text & "','" & Comboshift.Text & "','" & CInt(lblperday.Caption) & "')")
    End If
    
    con.Execute ("commit")
    
    MsgBox (" Employee ID is " & a)
    
End Sub

Private Sub cmdback_Click()
Unload frmemployee
frmmain.Show
End Sub

Private Sub cmdclear_Click()
    txtnm.Text = ""
    txtaddr.Text = ""
    ComboGender.Text = ""
    txtemail.Text = ""
    Combodes.Text = ""
    
    If (Combodes.Text = "Office Worker") Then
        txtsal.Text = ""
        combojob.Text = ""
    End If
End Sub

Private Sub cmddel_Click()
    
    Dim a As Integer

    a = InputBox(" Enter The Employee ID to Delete : ")
    res.Open "SELECT * from Employee where E_id = '" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    If res.BOF <> True Then
    con.Execute ("delete from Employee where E_id ='" & a & "'")
    con.Execute ("COMMIT")
    MsgBox ("Record Deleted Successfully..........")
    Else
    MsgBox ("RECORD NOT FOUND........")
    End If
    res.Close

End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdsearch_Click()
    Dim a As Integer
    
    txtnm.Text = ""
    ComboGender.Text = ""
    txtdob.Text = ""
    txtaddr.Text = ""
    txtemail.Text = ""
    Combodes.Text = ""
    
    a = InputBox("Enter The Employee ID to Search : ")

    res.Open "SELECT * from Employee where E_id = '" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    If res.BOF <> True Then
    txtnm.Text = res(1)
    txtdob.Text = res(2)
    ComboGender.Text = res(3)
    txtaddr.Text = res(4)
    txtemail.Text = res(5)
    Combodes.Text = res(6)
    
    MsgBox ("RECORD FOUND SUCCESSFULLY........")
    Else
    MsgBox ("RECORD NOT FOUND........")
    End If
    res.Close

End Sub

Private Sub cmdupdate_Click()
    Dim a As Integer
        
    a = InputBox("Enter The Employee ID to Update : ")
    
    res.Open "SELECT * from Employee where E_id = '" & a & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    If res.BOF <> True Then
    
        con.Execute ("update Employee set E_name='" & txtnm.Text & "',E_dob='" & txtdob.Text & "',E_gender='" & ComboGender.Text & "',E_add='" & txtaddr.Text & "',E_email='" & txtemail.Text & "','" & Combodes.Text & "' where A_id='" & a & "'")
        con.Execute ("commit")
    MsgBox ("Record Updated Successfully.....")
    Else
    MsgBox ("RECORD NOT FOUND........")
    End If
    res.Close

End Sub

Private Sub Combodes_Click()
    
    str = Combodes.Text
    
    If (str = "Office Worker") Then
        txtsal.Visible = True
        combojob.Visible = True
        lblsal.Visible = True
    lblpost.Visible = True
    
    Else
        txtsal.Visible = False
        combojob.Visible = False
      End If
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set res = New ADODB.Recordset
    
    con.Open "Provider=MSDAORA.1;Password=tiger;User ID=system;Persist Security Info=True;datasource=ADITYA"
    MsgBox ("Connection Established..........")
    
    txtsal.Visible = False
    combojob.Visible = False
    lblsal.Visible = False
    lblpost.Visible = False
    

End Sub

Private Sub txtmno_KeyPress(KeyAscii As Integer)
    Dim key As String
    
    key = Chr$(KeyAscii) 'ascii no to character string
    
    If ((key < "0" Or key > "9") And key <> ":") Then
        KeyAscii = 0
        Beep
    End If

End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii = Asc("@") Or KeyAscii = Asc("_") Or KeyAscii = Asc(".") Then
Else
    MsgBox ("INVALID INPU")
End If

End Sub

Private Sub txtnm_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub
