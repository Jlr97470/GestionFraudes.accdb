Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private StrWhere(3) As String

Private Function CreeListeDeclaration() As Boolean
   Dim StrSQL As String
   Dim StrSQLWhere As String

   On Error GoTo Err_CreeListeDeclaration

   StrSQL = "SELECT DecNumero, DecDate, DecDateCreation, DecDeclarantNom & ' ' & DecDeclarantPrenom AS [Nom Prenom Declarant], DecVictimeNom & ' ' & DecVictimePrenom AS [Nom Prenom Victime], DecBeneficiaireNom & ' ' & DecBeneficiairePrenom AS [Nom Prenom Beneficiaire], DecFraudeMontant FROM SelDeclaration " & _
                  "WHERE DecNumero In (SELECT DecNumero FROM TBLDECLARATION "

   Select Case StrWhere(1) <> "" Or StrWhere(2) <> "" Or StrWhere(3) <> ""
      Case True

                  StrSQL = StrSQL + "WHERE "

      Case False

   End Select

   Select Case StrWhere(1) = ""
      Case True

      Case False

         StrSQL = StrSQL + "DecDeclarantEntNumero=" + StrWhere(1)

         StrSQLWhere = " AND "

   End Select

   Select Case StrWhere(2) = ""
      Case True

      Case False

         StrSQL = StrSQL + StrSQLWhere + "DecNumero=" + StrWhere(2)

         StrSQLWhere = " AND "

   End Select

   Select Case StrWhere(3) = ""
      Case True

      Case False

         StrSQL = StrSQL + StrSQLWhere + "DecTypDecNumero=" + StrWhere(3)

   End Select

   StrSQL = StrSQL + ") ORDER BY DecNumero;"

   LstConsultationEntite.RowSource = StrSQL

   LstConsultationEntite.Requery

   Select Case LstConsultationEntite.ListCount
      Case 2

         CmbEntite = DLookup("DecDeclarantEntNumero", "TBLDECLARATION", "DecNumero=" & LstConsultationEntite.Column(0, 1))

      Case Else

   End Select

Exit_CreeListeDeclaration:

   Exit Function

Err_CreeListeDeclaration:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CreeListeDeclaration
End Function

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   CreeListeDeclaration

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_Form_Load
End Sub

Private Sub Form_Activate()

   On Error GoTo Err_Form_Activate

   DoCmd.Maximize

Exit_Form_Activate:

   Exit Sub

Err_Form_Activate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_Form_Activate
End Sub

Private Sub CmbEntite_AfterUpdate()

   On Error GoTo Err_CmbEntite_AfterUpdate

   StrWhere(1) = IIf(IsNull(CmbEntite.Column(0)), "", CmbEntite.Column(0))

Exit_CmbEntite_AfterUpdate:

   Exit Sub

Err_CmbEntite_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbEntite_AfterUpdate
End Sub

Private Sub TxtNumero_AfterUpdate()

   On Error GoTo Err_TxtNumero_Change

   StrWhere(2) = IIf(IsNull(TxtNumero), "", TxtNumero)

Exit_TxtNumero_Change:

   Exit Sub

Err_TxtNumero_Change:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtNumero_Change
End Sub

Private Sub CmbTypeDeclaration_AfterUpdate()

   On Error GoTo Err_CmbTypeDeclaration_AfterUpdate

   StrWhere(3) = IIf(IsNull(CmbTypeDeclaration.Column(0)), "", CmbTypeDeclaration.Column(0))

Exit_CmbTypeDeclaration_AfterUpdate:

   Exit Sub

Err_CmbTypeDeclaration_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbTypeDeclaration_AfterUpdate
End Sub

Private Sub LstConsultationEntite_DblClick(Cancel As Integer)

   On Error GoTo Err_LstConsultationEntite_DblClick

   DoCmd.OpenReport "EtaDeclaration", acViewPreview, , "DecNumero=" & LstConsultationEntite.Column(0), acWindowNormal

Exit_LstConsultationEntite_DblClick:

   Exit Sub

Err_LstConsultationEntite_DblClick:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_LstConsultationEntite_DblClick:
End Sub

Private Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   CreeListeDeclaration

   Select Case LstConsultationEntite.ListCount
      Case 0

         MsgBox "Aucunne déclaration correspondant au(x) critére(s) selectionné(s).", vbCritical

      Case Else

   End Select

Exit_CmdValider_Click:

   Exit Sub

Err_CmdValider_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdValider_Click
End Sub

Private Sub CmdQuitter_Click()

   On Error GoTo Err_CmdQuitter_Click

   DoCmd.Close acForm, "FrmConsultationEntite", acSavePrompt

Exit_CmdQuitter_Click:

   Exit Sub

Err_CmdQuitter_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdQuitter_Click
End Sub
