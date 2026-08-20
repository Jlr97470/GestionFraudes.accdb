Attribute VB_Name = "ModFonction"
Option Explicit

Public Function CreeMailDeclaration(ByVal StrExpediteur As String, ByVal StrDestinataire As String, ByVal StrSujet As String, ByVal StrBody As String, ByVal StrCritere As String) As Boolean
   Dim AppOutlook As New Outlook.Application
   Dim ObjMail As Outlook.MailItem
   Dim ObjRecipient As Outlook.Recipient
   Dim StrChamp As String
   Dim RsEnregistrement As DAO.Recordset

   On Error GoTo Err_CreeMailDeclaration

   CreeMailDeclaration = True

   Set RsEnregistrement = CurrentDb.OpenRecordset("SelDeclaration")

   RsEnregistrement.FindFirst StrCritere

   Set ObjMail = AppOutlook.CreateItem(olMailItem)

   Set ObjRecipient = ObjMail.Recipients.Add(StrDestinataire)

   With ObjMail

      .SentOnBehalfOfName = StrExpediteur

      .BodyFormat = olFormatRichText

      .Subject = StrSujet

      Do While InStr(StrBody, "[") > 0

         StrChamp = Mid(StrBody, InStr(StrBody, "[") + 1, InStr(InStr(StrBody, "["), StrBody, "]") - InStr(StrBody, "[") - 1)

         StrBody = Left(StrBody, InStr(StrBody, "[") - 1) + CStr(IIf(IsNull(RsEnregistrement(StrChamp)), vbNullString, RsEnregistrement(StrChamp))) + Mid(StrBody, InStr(StrBody, "[") + Len(StrChamp) + 2)

      Loop

      .Body = StrBody

   End With

   ObjMail.Display

Exit_CreeMailDeclaration:

   Exit Function

Err_CreeMailDeclaration:

   CreeMailDeclaration = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CreeMailDeclaration
End Function

Public Function CreeExcelConsultationCritere(ByRef StrCritere() As String, ByRef StrCritereValeur() As String, ByVal StrSQLDetail As String, ByVal StrSQLTotal As String) As Boolean
   Dim AppExcel As New Excel.Application
   Dim WkbClasseur As Excel.Workbook
   Dim WksFeuille As Excel.Worksheet
   Dim RsDetail As DAO.Recordset
   Dim RsTotal As DAO.Recordset
   Dim FldChamp As DAO.Field
   Dim IntColonne As Integer
   Dim IntLigne As Integer

   On Error GoTo Err_CreeExcelConsultationCritere

   CreeExcelConsultationCritere = True

   Set RsDetail = CurrentDb.OpenRecordset(StrSQLDetail)

   Set RsTotal = CurrentDb.OpenRecordset(StrSQLTotal)

   AppExcel.Visible = True

   Set WkbClasseur = AppExcel.Workbooks.Add

   Set WksFeuille = WkbClasseur.Worksheets(1)

   With WksFeuille

      .Name = "Consultation par critéres"

      IntLigne = 1

      .Cells(IntLigne, 1) = "CONSULTATION DES DECLARATIONS PAR CRITERES"

      .Range(.Cells(IntLigne, 1), .Cells(IntLigne, 5)).Merge

      IntLigne = 3

      .Cells(IntLigne, 1) = "Critéres:"

      For IntColonne = 2 To 5

         .Cells(IntLigne, IntColonne) = StrCritere(IntColonne - 1)

      Next

      IntLigne = 4

      .Cells(IntLigne, 1) = "Valeur:"

      For IntColonne = 2 To 5

         .Cells(IntLigne, IntColonne) = StrCritereValeur(IntColonne - 1)

      Next

      IntColonne = 1

      IntLigne = 6

      For Each FldChamp In RsDetail.Fields

            On Error Resume Next

            .Cells(IntLigne, IntColonne) = FldChamp.Properties("Description")

            Select Case Err.Number
               Case 0

               Case 3270

                  Err.Clear


               Case Else

                  GoTo Err_CreeExcelConsultationCritere

            End Select

            .Cells(IntLigne, IntColonne) = FldChamp.Name

            .Cells(IntLigne, IntColonne).Borders.LineStyle = 1

            .Cells(IntLigne, IntColonne).Borders.Weight = 3

            .Cells(IntLigne, IntColonne).HorizontalAlignment = -4108

            .Cells(IntLigne, IntColonne).Interior.Color = vbYellow

            On Error GoTo Err_CreeExcelConsultationCritere

            IntColonne = IntColonne + 1

      Next

      IntColonne = 1

      IntLigne = IntLigne + 1

      .Cells(IntLigne, IntColonne).CopyFromRecordset RsDetail

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsDetail.RecordCount - 1, IntColonne + RsDetail.Fields.Count - 1)).Borders.LineStyle = 1

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsDetail.RecordCount - 1, IntColonne + RsDetail.Fields.Count - 1)).Borders(8).Weight = 3

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsDetail.RecordCount - 1, IntColonne + RsDetail.Fields.Count - 1)).Borders(9).Weight = 3

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsDetail.RecordCount - 1, IntColonne + RsDetail.Fields.Count - 1)).Borders(7).Weight = 3

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsDetail.RecordCount - 1, IntColonne + RsDetail.Fields.Count - 1)).Borders(10).Weight = 3

      IntLigne = IntLigne + RsDetail.RecordCount

      .Cells(IntLigne, IntColonne).CopyFromRecordset RsTotal

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsTotal.RecordCount - 1, IntColonne + RsTotal.Fields.Count - 1)).Borders.LineStyle = 1

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsTotal.RecordCount - 1, IntColonne + RsTotal.Fields.Count - 1)).Borders(8).Weight = 3

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsTotal.RecordCount - 1, IntColonne + RsTotal.Fields.Count - 1)).Borders(9).Weight = 3

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsTotal.RecordCount - 1, IntColonne + RsTotal.Fields.Count - 1)).Borders(7).Weight = 3

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsTotal.RecordCount - 1, IntColonne + RsTotal.Fields.Count - 1)).Borders(10).Weight = 3

      .Range(.Cells(IntLigne, IntColonne), .Cells(IntLigne + RsTotal.RecordCount - 1, IntColonne + RsTotal.Fields.Count - 1)).Interior.Color = 12632256

      IntColonne = 1

      For Each FldChamp In RsDetail.Fields

            .Columns(IntColonne).AutoFit

            IntColonne = IntColonne + 1

      Next

      RsDetail.Close

      RsTotal.Close

   End With

Exit_CreeExcelConsultationCritere:

   Set RsDetail = Nothing

   Set RsTotal = Nothing

   Exit Function

Err_CreeExcelConsultationCritere:

   CreeExcelConsultationCritere = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CreeExcelConsultationCritere
End Function

Public Function CreeExcelRepartitionEntite() As Boolean
   Dim AppExcel As New Excel.Application
   Dim WkbClasseur As Excel.Workbook
   Dim WksFeuille As Excel.Worksheet
   Dim ChtGraphique As Excel.Chart
   Dim RsEntiteRepartitionDeclaration As DAO.Recordset
   Dim RsRepartitionDeclaration As DAO.Recordset

   On Error GoTo Err_CreeExcelRepartitionEntite

   CreeExcelRepartitionEntite = True

   Set RsEntiteRepartitionDeclaration = CurrentDb.OpenRecordset("SelEntiteRepartitionDeclaration")

   Set RsRepartitionDeclaration = CurrentDb.OpenRecordset("SelRepartitionDeclaration")

   AppExcel.Visible = True

   Set WkbClasseur = AppExcel.Workbooks.Open(CurrentProject.Path & "\ModeleRepartitionDeclaration.XLT")

   Set WksFeuille = WkbClasseur.Worksheets("Repartition")

   With WksFeuille

      .Cells(1, 7) = Date

      .Cells(7, 1).CopyFromRecordset RsEntiteRepartitionDeclaration

      .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 1).CopyFromRecordset RsRepartitionDeclaration

      .Range(.Cells(7, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount - 1, 10)).Borders.LineStyle = 1

      .Range(.Cells(7, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount - 1, 10)).Borders(8).Weight = 3

      .Range(.Cells(7, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount - 1, 10)).Borders(9).Weight = 3

      .Range(.Cells(7, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount - 1, 10)).Borders(7).Weight = 3

      .Range(.Cells(7, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount - 1, 10)).Borders(10).Weight = 3

      .Range(.Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 10)).Borders.LineStyle = 1

      .Range(.Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 10)).Borders(8).Weight = 3

      .Range(.Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 10)).Borders(9).Weight = 3

      .Range(.Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 10)).Borders(7).Weight = 3

      .Range(.Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 10)).Borders(10).Weight = 3

      .Range(.Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 1), .Cells(7 + RsEntiteRepartitionDeclaration.RecordCount, 10)).Interior.Color = 12632256

   End With

   Set ChtGraphique = WkbClasseur.Charts("Graphique Préjudice Net")

   ChtGraphique.SeriesCollection(1).FormulaLocal = "=SERIE(" & Chr(34) & "Préjudice Net" & Chr(34) & ";Repartition!$A$7:$A$" & (RsEntiteRepartitionDeclaration.RecordCount + 6) & ";Repartition!$F$7:$F$" & (RsEntiteRepartitionDeclaration.RecordCount + 6) & ";1)"

   ChtGraphique.SeriesCollection(2).FormulaLocal = "=SERIE(" & Chr(34) & "Nombre de fraudes" & Chr(34) & ";Repartition!$A$7:$A$" & (RsEntiteRepartitionDeclaration.RecordCount + 6) & ";Repartition!$B$7:$B$" & (RsEntiteRepartitionDeclaration.RecordCount + 6) & ";2)"

   Set ChtGraphique = WkbClasseur.Charts("Graphique Risque Global")

   ChtGraphique.SeriesCollection(1).FormulaLocal = "=SERIE(" & Chr(34) & "Nombre de fraudes" & Chr(34) & ";Repartition!$A$7:$A$" & (RsEntiteRepartitionDeclaration.RecordCount + 6) & ";Repartition!$K$7:$K$" & (RsEntiteRepartitionDeclaration.RecordCount + 6) & ";1)"

   ChtGraphique.SeriesCollection(2).FormulaLocal = "=SERIE(" & Chr(34) & "Montants des fraudes" & Chr(34) & ";Repartition!$A$7:$A$" & (RsEntiteRepartitionDeclaration.RecordCount + 6) & ";Repartition!$I$7:$I$" & (RsEntiteRepartitionDeclaration.RecordCount + 6) & ";2)"

   RsEntiteRepartitionDeclaration.Close

   RsRepartitionDeclaration.Close

Exit_CreeExcelRepartitionEntite:

   Set RsEntiteRepartitionDeclaration = Nothing

   Set RsRepartitionDeclaration = Nothing

   Exit Function

Err_CreeExcelRepartitionEntite:

   CreeExcelRepartitionEntite = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CreeExcelRepartitionEntite
End Function

Public Function CreeExcelRecapitulatifMois(ByVal IntType As Integer, ByVal DatDebut As Date, ByVal DatFin As Date) As Boolean
   Dim AppExcel As New Excel.Application
   Dim WkbClasseur As Excel.Workbook
   Dim WksFeuille As Excel.Worksheet
   Dim ChtGraphique As Excel.Chart
   Dim FldChamp As DAO.Field
   Dim QdfSQL As DAO.QueryDef
   Dim RsMoisRecapitulatif As DAO.Recordset
   Dim RsRecapitulatif As DAO.Recordset
   Dim RsRecapitulatifDeclaration As DAO.Recordset
   Dim RsRecapitulatifDeclarationMois As DAO.Recordset
   Dim IntColonne As Integer

   On Error GoTo Err_CreeExcelRepartitionEntite

   CreeExcelRecapitulatifMois = True

   Set QdfSQL = CurrentDb.QueryDefs("SelMoisRecapitulatif")

   QdfSQL.Parameters("Formulaires!FrmTableauBordStatistique!TxtDateDebut") = DatDebut

   QdfSQL.Parameters("Formulaires!FrmTableauBordStatistique!TxtDateFin") = DatFin

   Set RsMoisRecapitulatif = QdfSQL.OpenRecordset()

   Set QdfSQL = CurrentDb.QueryDefs("SelRecapitulatif")

   QdfSQL.Parameters("Formulaires!FrmTableauBordStatistique!TxtDateDebut") = DatDebut

   QdfSQL.Parameters("Formulaires!FrmTableauBordStatistique!TxtDateFin") = DatFin

   Set RsRecapitulatif = QdfSQL.OpenRecordset()

   AppExcel.Visible = True

   Set WkbClasseur = AppExcel.Workbooks.Open(CurrentProject.Path & "\ModeleRecapitulatifDeclaration.XLT")

   Set WksFeuille = WkbClasseur.Worksheets("Recapitulatif")

   With WksFeuille

      .Cells(1, 10) = Date

      .Cells(6, 1).CopyFromRecordset RsMoisRecapitulatif

      IntColonne = 1

      For Each FldChamp In RsRecapitulatif.Fields

         Select Case IntColonne
            Case 9, 13

               IntColonne = IntColonne + 1

            Case Else

         End Select

         .Cells(6 + RsMoisRecapitulatif.RecordCount, IntColonne) = FldChamp

         IntColonne = IntColonne + 1

      Next

      .Range(.Cells(6, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount - 1, 14)).Borders.LineStyle = 1

      .Range(.Cells(6, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount - 1, 14)).Borders(8).Weight = 3

      .Range(.Cells(6, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount - 1, 14)).Borders(9).Weight = 3

      .Range(.Cells(6, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount - 1, 14)).Borders(7).Weight = 3

      .Range(.Cells(6, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount - 1, 14)).Borders(10).Weight = 3

      .Range(.Cells(6 + RsMoisRecapitulatif.RecordCount, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount, 14)).Borders.LineStyle = 1

      .Range(.Cells(6 + RsMoisRecapitulatif.RecordCount, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount, 14)).Borders(8).Weight = 3

      .Range(.Cells(6 + RsMoisRecapitulatif.RecordCount, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount, 14)).Borders(9).Weight = 3

      .Range(.Cells(6 + RsMoisRecapitulatif.RecordCount, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount, 14)).Borders(7).Weight = 3

      .Range(.Cells(6 + RsMoisRecapitulatif.RecordCount, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount, 14)).Borders(10).Weight = 3

      .Range(.Cells(6 + RsMoisRecapitulatif.RecordCount, 1), .Cells(6 + RsMoisRecapitulatif.RecordCount, 14)).Interior.Color = 12632256

   End With

   Set ChtGraphique = WkbClasseur.Charts("Graphique Débit Fraude Avérées")

   ChtGraphique.SeriesCollection(1).FormulaLocal = "=SERIE(" & Chr(34) & "Nombre de fraudes" & Chr(34) & ";Recapitulatif!$A$6:$A$" & (RsMoisRecapitulatif.RecordCount + 6) & ";Recapitulatif!$B$6:$B$" & (RsMoisRecapitulatif.RecordCount + 6) & ";1)"

   ChtGraphique.SeriesCollection(2).FormulaLocal = "=SERIE(" & Chr(34) & "Montants des fraudes" & Chr(34) & ";Recapitulatif!$A$6:$A$" & (RsMoisRecapitulatif.RecordCount + 6) & ";Recapitulatif!$E$6:$E$" & (RsMoisRecapitulatif.RecordCount + 6) & ";2)"

   Select Case IntType
      Case 0

         AppExcel.DisplayAlerts = False

         ChtGraphique.Delete

         AppExcel.DisplayAlerts = True

      Case Else

   End Select

   Set ChtGraphique = WkbClasseur.Charts("Graphique Préjudice Net")

   ChtGraphique.SeriesCollection(1).FormulaLocal = "=SERIE(" & Chr(34) & "Nombre de fraudes" & Chr(34) & ";Recapitulatif!$A$6:$A$" & (RsMoisRecapitulatif.RecordCount + 6) & ";Recapitulatif!$B$6:$B$" & (RsMoisRecapitulatif.RecordCount + 6) & ";1)"

   ChtGraphique.SeriesCollection(2).FormulaLocal = "=SERIE(" & Chr(34) & "Préjudice Net" & Chr(34) & ";Recapitulatif!$A$6:$A$" & (RsMoisRecapitulatif.RecordCount + 6) & ";Recapitulatif!$J$6:$J$" & (RsMoisRecapitulatif.RecordCount + 6) & ";2)"

   Select Case IntType
      Case 1

         AppExcel.DisplayAlerts = False

         ChtGraphique.Delete

         AppExcel.DisplayAlerts = True

      Case Else

   End Select

   Set ChtGraphique = WkbClasseur.Charts("Graphique Risque Global")

   ChtGraphique.SeriesCollection(1).FormulaLocal = "=SERIE(" & Chr(34) & "Nombre de fraudes" & Chr(34) & ";Recapitulatif!$A$6:$A$" & (RsMoisRecapitulatif.RecordCount + 6) & ";Recapitulatif!$O$6:$O$" & (RsMoisRecapitulatif.RecordCount + 6) & ";1)"

   ChtGraphique.SeriesCollection(2).FormulaLocal = "=SERIE(" & Chr(34) & "Montants des fraudes" & Chr(34) & ";Recapitulatif!$A$6:$A$" & (RsMoisRecapitulatif.RecordCount + 6) & ";Recapitulatif!$N$6:$N$" & (RsMoisRecapitulatif.RecordCount + 6) & ";2)"

   Select Case IntType
      Case 1

         AppExcel.DisplayAlerts = False

         ChtGraphique.Delete

         AppExcel.DisplayAlerts = True

      Case Else

   End Select

   RsMoisRecapitulatif.MoveFirst

   Set QdfSQL = CurrentDb.QueryDefs("SelRecapitulatifDeclaration")

   QdfSQL.Parameters("Formulaires!FrmTableauBordStatistique!TxtDateDebut") = DatDebut

   QdfSQL.Parameters("Formulaires!FrmTableauBordStatistique!TxtDateFin") = DatFin

   Set RsRecapitulatifDeclaration = QdfSQL.OpenRecordset()

   Do

      Set WksFeuille = WkbClasseur.Worksheets("Mois")

      WksFeuille.Copy After:=WksFeuille

      Set WksFeuille = WkbClasseur.Worksheets("Mois (2)")

      With WksFeuille

         .Name = RsMoisRecapitulatif("FraudeMois")

         .Cells(5, 4) = RsMoisRecapitulatif("FraudeMois")

         .Cells(7, 4) = RsMoisRecapitulatif("NbAverees")

         .Cells(8, 4) = RsMoisRecapitulatif("NbCreditAverees")

         .Cells(9, 4) = RsMoisRecapitulatif("NbDebitAverees")

         .Cells(10, 4) = RsMoisRecapitulatif("NbEntiteAverees")

         .Cells(11, 4) = RsMoisRecapitulatif("MontantDebitAverees")

         .Cells(12, 4) = RsMoisRecapitulatif("MontantCreditAverees")

         .Cells(13, 4) = RsMoisRecapitulatif("MoyenneCreditAverees")

         .Cells(14, 4) = RsMoisRecapitulatif("MoyenneDebitAverees")

         RsRecapitulatifDeclaration.Filter = "MonthName(Month([DecFraudeDate])) & ' ' & Year([DecFraudeDate])='" & RsMoisRecapitulatif("FraudeMois") & "'"

         Set RsRecapitulatifDeclarationMois = RsRecapitulatifDeclaration.OpenRecordset()

         .Cells(20, 1).CopyFromRecordset RsRecapitulatifDeclarationMois

         .Range(.Cells(20, 1), .Cells(20 + RsRecapitulatifDeclarationMois.RecordCount - 1, 7)).Borders.LineStyle = 1

         .Range(.Cells(20, 1), .Cells(20 + RsRecapitulatifDeclarationMois.RecordCount - 1, 7)).Borders(8).Weight = 3

         .Range(.Cells(20, 1), .Cells(20 + RsRecapitulatifDeclarationMois.RecordCount - 1, 7)).Borders(9).Weight = 3

         .Range(.Cells(20, 1), .Cells(20 + RsRecapitulatifDeclarationMois.RecordCount - 1, 7)).Borders(7).Weight = 3

         .Range(.Cells(20, 1), .Cells(20 + RsRecapitulatifDeclarationMois.RecordCount - 1, 7)).Borders(10).Weight = 3

         RsRecapitulatifDeclarationMois.Close

      End With

      RsMoisRecapitulatif.MoveNext

   Loop Until RsMoisRecapitulatif.EOF = True

   Set WksFeuille = WkbClasseur.Worksheets("Mois")

   AppExcel.DisplayAlerts = False

   WksFeuille.Delete

   AppExcel.DisplayAlerts = False

   RsMoisRecapitulatif.Close

   RsRecapitulatif.Close

Exit_CreeExcelRepartitionEntite:

   Set RsMoisRecapitulatif = Nothing

   Set RsRecapitulatif = Nothing

   Exit Function

Err_CreeExcelRepartitionEntite:

   CreeExcelRecapitulatifMois = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CreeExcelRepartitionEntite
End Function
