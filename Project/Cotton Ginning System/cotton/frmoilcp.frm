VERSION 5.00
Begin VB.Form frmoilcp 
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8055
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
      Width           =   14655
      Begin VB.CommandButton cmdsubmit 
         Caption         =   "SUBMIT"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   6480
         Width           =   2415
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
         Height          =   1095
         Left            =   4440
         TabIndex        =   9
         Top             =   6480
         Width           =   2415
      End
      Begin VB.CommandButton cmdprevious 
         Caption         =   "PREVIOUS"
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
         Left            =   7680
         TabIndex        =   8
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0084ECEC&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5415
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   12615
         Begin VB.ComboBox cmbtype 
            Height          =   315
            ItemData        =   "frmoilcp.frx":0000
            Left            =   5280
            List            =   "frmoilcp.frx":000A
            TabIndex        =   2
            Top             =   1680
            Width           =   3135
         End
         Begin VB.Label lbladd1 
            Height          =   1335
            Left            =   5280
            TabIndex        =   13
            Top             =   3120
            Width           =   3255
         End
         Begin VB.Label lblmno 
            Height          =   375
            Left            =   5280
            TabIndex        =   12
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label lblinm 
            Height          =   375
            Left            =   5280
            TabIndex        =   11
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label lblnm 
            Height          =   375
            Left            =   5280
            TabIndex        =   10
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label lblid 
            BackStyle       =   0  'Transparent
            Caption         =   "EXPORT  ID"
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
            Left            =   720
            TabIndex        =   7
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblname 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPANY NAME"
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
            Left            =   720
            TabIndex        =   6
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label lbltype 
            BackStyle       =   0  'Transparent
            Caption         =   "TYPE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   5
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblcp 
            BackStyle       =   0  'Transparent
            Caption         =   "CONTCT  NO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   4
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label lbladd 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPANY ADDRESS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   600
            TabIndex        =   3
            Top             =   3360
            Width           =   2775
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   11775
      Left            =   -120
      Picture         =   "frmoilcp.frx":0029
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20415
   End
End
Attribute VB_Name = "frmoilcp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
