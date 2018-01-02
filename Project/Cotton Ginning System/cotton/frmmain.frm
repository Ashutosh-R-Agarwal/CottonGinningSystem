VERSION 5.00
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10650
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11400
   LinkTopic       =   "MAINFORM"
   Picture         =   "frmmain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuexit 
      Caption         =   "&EXIT"
   End
   Begin VB.Menu mnureg 
      Caption         =   "&REGISTRATION"
      Begin VB.Menu mnuagent 
         Caption         =   "AGENT"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuemp 
         Caption         =   "EMPLOYEE"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnucomp 
         Caption         =   "COMPANY"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnubill 
      Caption         =   "&BILLING"
      Begin VB.Menu mnupur 
         Caption         =   "&PURCHASE ORDER"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuinvoicereport 
         Caption         =   "INVOICE"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuds 
      Caption         =   "&DAILY SCHEDULE"
      Begin VB.Menu mnuos 
         Caption         =   "ONSITE SCHEDULE"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnupd 
         Caption         =   "&PROCESSING DETAILS"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuagnetsupp 
         Caption         =   "AGENT SUPPLIES"
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnusuppto 
         Caption         =   "SUPPLIED TO"
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnubillreport 
      Caption         =   "&BILING REPORTS"
      Begin VB.Menu mnupurchase 
         Caption         =   "&PURCHASE REPORT"
      End
      Begin VB.Menu mnuinvoice 
         Caption         =   "&BILL REPORT"
      End
   End
   Begin VB.Menu mnurep 
      Caption         =   "&REPORTS"
      Begin VB.Menu mnusupto 
         Caption         =   "&MONTHLy SUPLY TO"
      End
      Begin VB.Menu mnusepa 
         Caption         =   "-"
      End
      Begin VB.Menu mnusupplies 
         Caption         =   "&MOTHLY SUPPLIES"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuraw 
         Caption         =   "&RAWCOTTON STOCK"
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnustock 
         Caption         =   "&STOCK REPORT"
      End
      Begin VB.Menu mnulint 
         Caption         =   "&LINT  COMPANY REPORT"
      End
      Begin VB.Menu mnuoil 
         Caption         =   "&OIL COMPANY REPORT"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuagent_Click()
frmagent.Show
End Sub

Private Sub mnuagnetsupp_Click()
frmsupplies.Show
End Sub

Private Sub mnucomp_Click()
frmexport.Show
End Sub

Private Sub mnuemp_Click()
frmemployee.Show
End Sub

Private Sub mnuex_Click()
frmexport.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuinvoice_Click()

Unload denpurchase
Unload drebill
drebill.Show
End Sub

Private Sub mnuinvoicereport_Click()
frminvoice.Show

End Sub

Private Sub mnulint_Click()
Unload denpurchase
Unload drplint
drplint.Show

End Sub

Private Sub mnuos_Click()
frmonsiteworker.Show
End Sub

Private Sub mnupd_Click()
frmprocess.Show
End Sub

Private Sub mnupur_Click()
frmpurchase.Show
End Sub

Private Sub mnusuppliedto_Click()
End Sub

Private Sub mnupurchase_Click()
Unload denpurchase
Unload drepurchase
drepurchase.Show

End Sub

Private Sub mnuraw_Click()
 Unload denpurchase
Unload drpraw
drpraw.Show

End Sub

Private Sub mnustock_Click()
Unload denpurchase
Unload drpstock
drpstock.Show

End Sub

Private Sub mnusupplies_Click()
Unload denpurchase
Unload dresupples
dresupples.Show

End Sub

Private Sub mnusuppto_Click()
frmsupplyto.Show
End Sub

Private Sub mnusupto_Click()
Unload denpurchase
Unload dreSupto
dreSupto.Show


End Sub
