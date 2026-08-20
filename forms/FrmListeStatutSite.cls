Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

   On Error GoTo Err_Form_Activate

   DoCmd.Maximize

Exit_Form_Activate:

   Exit Sub

Err_Form_Activate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_Form_Activate
End Sub

Private Sub CmdQuitter_Click()

   On Error GoTo Err_CmdQuitter_Click

   DoCmd.Close acForm, "FrmListeStatutSite", acSavePrompt

Exit_CmdQuitter_Click:

   Exit Sub

Err_CmdQuitter_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdQuitter_Click
End Sub
