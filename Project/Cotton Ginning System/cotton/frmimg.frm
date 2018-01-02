VERSION 5.00
Begin VB.Form frmimg 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   14880
   Begin VB.Image Image1 
      Height          =   9615
      Left            =   360
      Picture         =   "frmimg.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   20025
   End
End
Attribute VB_Name = "frmimg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmimg.Width = 20000
frmimg.Height = 15000
End Sub
