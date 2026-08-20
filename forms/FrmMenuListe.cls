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

Private Sub CmdListeTypeDeclaration_Click()

   On Error GoTo Err_CmdListeTypeDeclaration_Click

   DoCmd.OpenForm "FrmListeTypeDeclaration", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeTypeDeclaration_Click:

    Exit Sub

Err_CmdListeTypeDeclaration_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeTypeDeclaration_Click
End Sub

Private Sub CmdListeEntite_Click()

   On Error GoTo Err_CmdListeEntite_Click

   DoCmd.OpenForm "FrmListeEntite", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeEntite_Click:

    Exit Sub

Err_CmdListeEntite_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeEntite_Click
End Sub

Private Sub CmdListeMarche_Click()

   On Error GoTo Err_CmdListeMarche_Click

   DoCmd.OpenForm "FrmListeMarche", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeMarche_Click:

    Exit Sub

Err_CmdListeMarche_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeMarche_Click
End Sub

Private Sub CmdListeTypeCompte_Click()

   On Error GoTo Err_CmdListeTypeCompte_Click

   DoCmd.OpenForm "FrmListeTypeCompte", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeTypeCompte_Click:

    Exit Sub

Err_CmdListeTypeCompte_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeTypeCompte_Click
End Sub

Private Sub CmdTypeBeneficiaire_Click()

   On Error GoTo Err_CmdTypeBeneficiaire_Click

   DoCmd.OpenForm "FrmListeTypeBeneficiaire", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdTypeBeneficiaire_Click:

    Exit Sub

Err_CmdTypeBeneficiaire_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdTypeBeneficiaire_Click
End Sub

Private Sub CmdListeTypeOperation_Click()

   On Error GoTo Err_CmdListeTypeOperation_Click

   DoCmd.OpenForm "FrmListeTypeOperation", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeTypeOperation_Click:

    Exit Sub

Err_CmdListeTypeOperation_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeTypeOperation_Click
End Sub

Private Sub CmdListeTypeFraude_Click()

   On Error GoTo Err_CmdListeTypeFraude_Click

   DoCmd.OpenForm "FrmListeTypeFraude", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeTypeFraude_Click:

    Exit Sub

Err_CmdListeTypeFraude_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeTypeFraude_Click
End Sub

Private Sub CmdListeSourceIP_Click()

   On Error GoTo Err_CmdListeSourceIP_Click

   DoCmd.OpenForm "FrmListeSourceIP", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeSourceIP_Click:

    Exit Sub

Err_CmdListeSourceIP_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeSourceIP_Click
End Sub

Private Sub CmdListeAutorite_Click()

   On Error GoTo Err_CmdListeAutorite_Click

   DoCmd.OpenForm "FrmListeAutorite", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeAutorite_Click:

    Exit Sub

Err_CmdListeAutorite_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeAutorite_Click
End Sub

Private Sub CmdListeStatutSite_Click()

   On Error GoTo Err_CmdListeStatutSite_Click

   DoCmd.OpenForm "FrmListeStatutSite", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeStatutSite_Click:

    Exit Sub

Err_CmdListeStatutSite_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeStatutSite_Click
End Sub

Private Sub CmdListeSourceSite_Click()

   On Error GoTo Err_CmdListeSourceSite_Click

   DoCmd.OpenForm "FrmListeSourceSite", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeSourceSite_Click:

    Exit Sub

Err_CmdListeSourceSite_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeSourceSite_Click
End Sub

Private Sub CmdListeLocalisationSite_Click()

   On Error GoTo Err_CmdListeLocalisationSite_Click

   DoCmd.OpenForm "FrmListeLocalisationSite", acNormal, , , acFormEdit, acWindowNormal

Exit_CmdListeLocalisationSite_Click:

    Exit Sub

Err_CmdListeLocalisationSite_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdListeLocalisationSite_Click
End Sub

Private Sub CmdQuitter_Click()

   On Error GoTo Err_CmdQuitter_Click

   DoCmd.Close acForm, "FrmMenuListe", acSavePrompt

Exit_CmdQuitter_Click:

   Exit Sub

Err_CmdQuitter_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdQuitter_Click
End Sub
