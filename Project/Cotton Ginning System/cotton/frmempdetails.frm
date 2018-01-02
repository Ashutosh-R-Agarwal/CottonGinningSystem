VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmempdetails 
   Caption         =   "Form1"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   15120
   Begin VB.Frame Frame1 
      BackColor       =   &H00E4BAC4&
      Caption         =   "EMPLOYEE"
      Height          =   6495
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   13095
      Begin VB.CommandButton cmddel 
         Caption         =   "DELETE"
         Height          =   555
         Left            =   5280
         TabIndex        =   4
         Top             =   4800
         Width           =   1695
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         Height          =   555
         Left            =   3240
         TabIndex        =   3
         Top             =   4800
         Width           =   1695
      End
      Begin VB.CommandButton cmdmodify 
         Caption         =   "MODIFY"
         Height          =   555
         Left            =   1200
         TabIndex        =   2
         Top             =   4800
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
         _Version        =   393216
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
   End
   Begin VB.ComboBox cmbsearch 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblsearch 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH BY"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmempdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
