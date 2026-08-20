Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CreeListeCritere() As Boolean
   Dim RsDeclaration As DAO.Recordset
   Dim FldChamp As DAO.Field
   Dim IntNumero As Integer

   On Error GoTo Err_CreeListeCritere

   CreeListeCritere = True

   Set RsDeclaration = CurrentDb.OpenRecordset("SelDeclaration")

   For Each FldChamp In RsDeclaration.Fields

         For IntNumero = 1 To 4

            Me.Controls("CmbCritere" & IntNumero).AddItem FldChamp.Name & ";" & FldChamp.Properties("Description") & ";" & FldChamp.Type

         Next

   Next

   RsDeclaration.Close

Exit_CreeListeCritere:

   Set RsDeclaration = Nothing

   Exit Function

Err_CreeListeCritere:

   CreeListeCritere = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CreeListeCritere
End Function

Private Function CreeListeDeclaration() As Boolean
   Dim IntNumero As Integer
   Dim StrSQLDetail As String
   Dim StrSQLTotal As String
   Dim StrSQLDetailWhere As String
   Dim StrSQLTotalWhere As String

   On Error GoTo Err_CreeListeDeclaration

   StrSQLDetail = "TRANSFORM IIf(IsNull(Count([DecNumero])),0,Count([DecNumero])) AS NbDeclarationType " & _
                           "SELECT SelDeclaration.DecDeclarantEntNom AS Entite, Count(SelDeclaration.DecNumero) AS TOTAL FROM SelDeclaration "

   StrSQLTotal = "TRANSFORM Count(SelDeclaration.DecNumero) AS CompteDeDecNumero " & _
                         "SELECT 'TOTAL' AS [TOTAL GENERAL], Count(SelDeclaration.DecNumero) AS TOTAL FROM SelDeclaration "

   For IntNumero = 1 To 4

      Select Case Me.Controls("CmbCritere" & IntNumero).Column(0) <> "" And IIf(IsNull(Me.Controls("CmbCritere" & IntNumero & "Valeur")), vbNullString, Me.Controls("CmbCritere" & IntNumero & "Valeur")) <> ""
         Case True

               Select Case Me.Controls("CmbCritere" & IntNumero).Column(2)
                  Case dbText, dbMemo, dbChar, dbDate, dbTime, dbTimeStamp

                     StrSQLDetailWhere = StrSQLDetailWhere & Me.Controls("CmbCritere" & IntNumero).Column(0) & " LIKE " & Chr$(34) & "*" & Me.Controls("CmbCritere" & IntNumero & "Valeur") & "*" & Chr$(34) & " AND "

                     StrSQLTotalWhere = StrSQLTotalWhere & Me.Controls("CmbCritere" & IntNumero).Column(0) & " LIKE " & Chr$(34) & "*" & Me.Controls("CmbCritere" & IntNumero & "Valeur") & "*" & Chr$(34) & " AND "

                  Case dbBigInt, dbBinary, dbByte, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger, dbLong, dbLongBinary, dbNumeric, dbSingle, dbVarBinary

                     Select Case IsNumeric(Me.Controls("CmbCritere" & IntNumero & "Valeur"))
                        Case True

                           StrSQLDetailWhere = StrSQLDetailWhere & Me.Controls("CmbCritere" & IntNumero).Column(0) & "=" & Me.Controls("CmbCritere" & IntNumero & "Valeur") & " AND "

                           StrSQLTotalWhere = StrSQLTotalWhere & Me.Controls("CmbCritere" & IntNumero).Column(0) & "=" & Me.Controls("CmbCritere" & IntNumero & "Valeur") & " AND "

                        Case False

                     End Select

                  Case dbBoolean

                     Select Case Me.Controls("CmbCritere" & IntNumero & "Valeur")
                        Case "True", "Oui", "Vrai", "-1"

                           StrSQLDetailWhere = StrSQLDetailWhere & Me.Controls("CmbCritere" & IntNumero).Column(0) & "=True AND "

                           StrSQLTotalWhere = StrSQLTotalWhere & Me.Controls("CmbCritere" & IntNumero).Column(0) & "=True AND "

                        Case "False", "Non", "Faux", "0"

                           StrSQLDetailWhere = StrSQLDetailWhere & Me.Controls("CmbCritere" & IntNumero).Column(0) & "=False AND "

                           StrSQLTotalWhere = StrSQLTotalWhere & Me.Controls("CmbCritere" & IntNumero).Column(0) & "=False AND "

                     End Select

               End Select

         Case False

      End Select

   Next

   Select Case StrSQLDetailWhere
      Case vbNullString

      Case Else

         StrSQLDetail = StrSQLDetail & " WHERE " & Left(StrSQLDetailWhere, Len(StrSQLDetailWhere) - 4)

   End Select

   Select Case StrSQLTotalWhere
      Case vbNullString

      Case Else

         StrSQLTotal = StrSQLTotal & " WHERE " & Left(StrSQLTotalWhere, Len(StrSQLTotalWhere) - 4)

   End Select

   StrSQLDetail = StrSQLDetail + "GROUP BY SelDeclaration.DecDeclarantEntNom PIVOT SelDeclaration.DecTypDecNom;"

   StrSQLTotal = StrSQLTotal + "GROUP BY 'TOTAL' PIVOT SelDeclaration.DecTypDecNom;"

   LstConsultationCritereDetail.RowSource = StrSQLDetail

   LstConsultationCritereDetail.Requery

   LstConsultationCritereTotal.RowSource = StrSQLTotal

   LstConsultationCritereTotal.Requery

Exit_CreeListeDeclaration:

   Exit Function

Err_CreeListeDeclaration:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CreeListeDeclaration
End Function

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   CreeListeCritere

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

Private Sub CmbCritere1_AfterUpdate()
   Dim StrSQL As String

   On Error GoTo Err_CmbCritere1_AfterUpdate

   Select Case CmbCritere1.Column(2)
      Case dbBoolean

         CmbCritere1Valeur.RowSourceType = "Value List"

         StrSQL = "Oui;Non"

      Case Else

         CmbCritere1Valeur.RowSourceType = "Table/Query"

         StrSQL = "SELECT CStr(" & CmbCritere1.Column(0) & ") FROM SelDeclaration WHERE Not IsNull(" & CmbCritere1.Column(0) & ") GROUP BY CStr(" & CmbCritere1.Column(0) & "), " & CmbCritere1.Column(0) & " ORDER BY " & CmbCritere1.Column(0) & ";"

   End Select

   CmbCritere1Valeur.RowSource = StrSQL

   CmbCritere1Valeur.Requery

Exit_CmbCritere1_AfterUpdate:

   Exit Sub

Err_CmbCritere1_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbCritere1_AfterUpdate
End Sub

Private Sub CmbCritere2_AfterUpdate()
   Dim StrSQL As String

   On Error GoTo Err_CmbCritere2_AfterUpdate

   Select Case CmbCritere2.Column(2)
      Case dbBoolean

         CmbCritere1Valeur.RowSourceType = "Value List"

         StrSQL = "Oui;Non"

      Case Else

         CmbCritere2Valeur.RowSourceType = "Table/Query"

         StrSQL = "SELECT CStr(" & CmbCritere2.Column(0) & ") FROM SelDeclaration WHERE Not IsNull(" & CmbCritere2.Column(0) & ") GROUP BY CStr(" & CmbCritere2.Column(0) & "), " & CmbCritere2.Column(0) & " ORDER BY " & CmbCritere2.Column(0) & ";"

   End Select

   CmbCritere2Valeur.RowSource = StrSQL

   CmbCritere2Valeur.Requery

Exit_CmbCritere2_AfterUpdate:

   Exit Sub

Err_CmbCritere2_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbCritere2_AfterUpdate
End Sub

Private Sub CmbCritere3_AfterUpdate()
   Dim StrSQL As String

   On Error GoTo Err_CmbCritere3_AfterUpdate

   Select Case CmbCritere3.Column(2)
      Case dbBoolean

         CmbCritere3Valeur.RowSourceType = "Value List"

         StrSQL = "Oui;Non"

      Case Else

         CmbCritere3Valeur.RowSourceType = "Table/Query"

         StrSQL = "SELECT CStr(" & CmbCritere3.Column(0) & ") FROM SelDeclaration WHERE Not IsNull(" & CmbCritere3.Column(0) & ") GROUP BY CStr(" & CmbCritere3.Column(0) & "), " & CmbCritere3.Column(0) & " ORDER BY " & CmbCritere3.Column(0) & ";"

   End Select

   CmbCritere3Valeur.RowSource = StrSQL

   CmbCritere3Valeur.Requery

Exit_CmbCritere3_AfterUpdate:

   Exit Sub

Err_CmbCritere3_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbCritere3_AfterUpdate
End Sub

Private Sub CmbCritere4_AfterUpdate()
   Dim StrSQL As String

   On Error GoTo Err_CmbCritere4_AfterUpdate

   Select Case CmbCritere4.Column(2)
      Case dbBoolean

         CmbCritere4Valeur.RowSourceType = "Value List"

         StrSQL = "Oui;Non"

      Case Else

         CmbCritere4Valeur.RowSourceType = "Table/Query"

         StrSQL = "SELECT CStr(" & CmbCritere4.Column(0) & ") FROM SelDeclaration WHERE Not IsNull(" & CmbCritere4.Column(0) & ") GROUP BY CStr(" & CmbCritere4.Column(0) & "), " & CmbCritere4.Column(0) & " ORDER BY " & CmbCritere4.Column(0) & ";"

   End Select

   CmbCritere4Valeur.RowSource = StrSQL

   CmbCritere4Valeur.Requery

Exit_CmbCritere4_AfterUpdate:

   Exit Sub

Err_CmbCritere4_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbCritere4_AfterUpdate
End Sub

Private Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   CreeListeDeclaration

   Select Case LstConsultationCritereDetail.ListCount
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

   DoCmd.Close acForm, "FrmConsultationCritere", acSavePrompt

Exit_CmdQuitter_Click:

   Exit Sub

Err_CmdQuitter_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdQuitter_Click
End Sub

Private Sub CmdExporter_Click()
   Dim StrCritere(4) As String
   Dim StrValeur(4) As String

   On Error GoTo Err_CmdExporter_Click

   StrCritere(1) = IIf(IsNull(CmbCritere1.Column(1)), vbNullString, CmbCritere1.Column(1))

   StrCritere(2) = IIf(IsNull(CmbCritere2.Column(1)), vbNullString, CmbCritere2.Column(1))

   StrCritere(3) = IIf(IsNull(CmbCritere3.Column(1)), vbNullString, CmbCritere3.Column(1))

   StrCritere(4) = IIf(IsNull(CmbCritere4.Column(1)), vbNullString, CmbCritere4.Column(1))

   StrValeur(1) = IIf(IsNull(CmbCritere1Valeur), vbNullString, CmbCritere1Valeur)

   StrValeur(2) = IIf(IsNull(CmbCritere2Valeur), vbNullString, CmbCritere2Valeur)

   StrValeur(3) = IIf(IsNull(CmbCritere3Valeur), vbNullString, CmbCritere3Valeur)

   StrValeur(4) = IIf(IsNull(CmbCritere4Valeur), vbNullString, CmbCritere4Valeur)

   CreeExcelConsultationCritere StrCritere, StrValeur, LstConsultationCritereDetail.RowSource, LstConsultationCritereTotal.RowSource

Exit_CmdExporter_Click:

    Exit Sub

Err_CmdExporter_Click:

    MsgBox Err.Number & " " & Err.Description

    Resume Exit_CmdExporter_Click
End Sub
