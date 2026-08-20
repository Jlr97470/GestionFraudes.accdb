Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdExporter_Click()

   On Error GoTo Err_CmdExporter_Click

   CreeExcelRecapitulatifMois 0, TxtDateDebut, TxtDateFin

   DoCmd.Close acForm, "FrmTableauBordStatistiqueConsolidee", acSavePrompt

Exit_CmdExporter_Click:

   Exit Sub

Err_CmdExporter_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdExporter_Click
End Sub
