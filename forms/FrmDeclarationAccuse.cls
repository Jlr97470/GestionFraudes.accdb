Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LngPositionText As String

Private Function CreeListeChamp() As Boolean
   Dim RsDeclaration As DAO.Recordset
   Dim FldChamp As DAO.Field

   On Error GoTo Err_CreeListeChamp

   CreeListeChamp = True

   Set RsDeclaration = CurrentDb.OpenRecordset("SelDeclaration")

   For Each FldChamp In RsDeclaration.Fields

      CmbChamp.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description")

   Next

   RsDeclaration.Close

Exit_CreeListeChamp:

   Set RsDeclaration = Nothing

   Exit Function

Err_CreeListeChamp:

   CreeListeChamp = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CreeListeChamp
End Function

Private Sub Form_Load()
   Dim RsParametre As DAO.Recordset

   On Error GoTo Err_Form_Load

   Set RsParametre = CurrentDb.OpenRecordset("SELECT TBLPARAMETRE.* FROM TBLPARAMETRE WHERE ParCode='TXTMAI';")

   TxtAccuse = RsParametre("ParValeur")

   RsParametre.Close

   CreeListeChamp

Exit_Form_Load:

   Set RsParametre = Nothing

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

Private Sub Form_Unload(Cancel As Integer)
   Dim RsParametre As DAO.Recordset

   On Error GoTo Err_Form_Unload

   Set RsParametre = CurrentDb.OpenRecordset("SELECT TBLPARAMETRE.* FROM TBLPARAMETRE WHERE ParCode='TXTMAI';")

   RsParametre.Edit

   RsParametre("ParValeur") = TxtAccuse

   RsParametre.Update

   RsParametre.Close

Exit_Form_Unload:

   Set RsParametre = Nothing

   Exit Sub

Err_Form_Unload:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_Form_Unload
End Sub

Private Sub TxtAccuse_Exit(Cancel As Integer)

   On Error GoTo Err_TxtAccuse_Exit

   LngPositionText = TxtAccuse.SelStart

Exit_TxtAccuse_Exit:

   Exit Sub

Err_TxtAccuse_Exit:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtAccuse_Exit

End Sub

Private Sub CmdInserer_Click()

   On Error GoTo Err_CmdInserer_Click

   TxtAccuse = Left(TxtAccuse, LngPositionText) & "[" & CmbChamp.Column(0) & "]" & Mid(TxtAccuse, LngPositionText + 1)

Exit_CmdInserer_Click:

   Exit Sub

Err_CmdInserer_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdInserer_Click
End Sub

Private Sub CmdQuitter_Click()

   On Error GoTo Err_CmdQuitter_Click

   DoCmd.Close acForm, "FrmDeclarationAccuse", acSavePrompt

Exit_CmdQuitter_Click:

   Exit Sub

Err_CmdQuitter_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdQuitter_Click
End Sub
