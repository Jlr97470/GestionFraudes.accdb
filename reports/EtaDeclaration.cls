Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Report_Activate()

   On Error GoTo Err_Report_Activate

   DoCmd.Maximize

Exit_Report_Activate:

   Exit Sub

Err_Report_Activate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_Report_Activate
End Sub

Private Sub Report_Open(Cancel As Integer)

   On Error GoTo Err_Report_Open

   DoCmd.Maximize

Exit_Report_Open:

   Exit Sub

Err_Report_Open:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_Report_Open
End Sub
