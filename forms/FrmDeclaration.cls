Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IntMode As Integer
Public IntModeOld As Integer
Public IntModeGeneral As Integer

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   ChangeMode Me, MODEDEPLACER

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_Form_Open
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

Private Sub Form_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_Form_BeforeUpdate

   Select Case IntMode
      Case MODEVALIDER

         Select Case IsNull(TxtDecDate)
            Case True

               MsgBox "Le champ date de déclaration  est obligatoire", , Me.Caption

               TxtDecDate.SetFocus

               Cancel = True

               GoTo Exit_Form_BeforeUpdate

            Case False

         End Select

         Select Case IsNull(CmbDecTypDecNumero)
            Case True

               MsgBox "Le champ type de déclaration est obligatoire", , Me.Caption

               CmbDecTypDecNumero.SetFocus

               Cancel = True

               GoTo Exit_Form_BeforeUpdate

            Case False

         End Select

         Select Case IsNull(CmbDecDeclarantEntNumero)
            Case True

               MsgBox "Le champ entité déclarant est obligatoire", , Me.Caption

               CmbDecDeclarantEntNumero.SetFocus

               Cancel = True

               GoTo Exit_Form_BeforeUpdate

            Case False

         End Select

         Select Case IsNull(TxtDecDeclarantNom)
            Case True

               MsgBox "Le champ nom déclarant est obligatoire", , Me.Caption

               TxtDecDeclarantNom.SetFocus

               Cancel = True

               GoTo Exit_Form_BeforeUpdate

            Case False

         End Select

         Select Case IsNull(TxtDecDeclarantTelephone1)
            Case True

               MsgBox "Le champ telephone 1 déclarant est obligatoire", , Me.Caption

               TxtDecDeclarantTelephone1.SetFocus

               Cancel = True

               GoTo Exit_Form_BeforeUpdate

            Case False

         End Select

         Select Case IsNull(TxtDecDeclarantMail)
            Case True

               MsgBox "Le champ mail déclarant est obligatoire", , Me.Caption

               TxtDecDeclarantMail.SetFocus

               Cancel = True

               GoTo Exit_Form_BeforeUpdate

            Case False

         End Select

         Select Case Me.DecTypDecNumero
            Case 3

               Select Case IsNull(TxtDecSiteIlliciteDate)
                  Case True

                       MsgBox "Le champ date de détection site illicite est obligatoire", , Me.Caption

                       TxtDecSiteIlliciteDate.SetFocus

                       Cancel = True

                       GoTo Exit_Form_BeforeUpdate

                  Case False

               End Select

               Select Case IsNull(CmbDecSiteIlliciteStaSitNumero)
                  Case True

                       MsgBox "Le champ statut site illicite est obligatoire", , Me.Caption

                       CmbDecSiteIlliciteStaSitNumero.SetFocus

                       Cancel = True

                       GoTo Exit_Form_BeforeUpdate

                  Case False

               End Select

               Select Case IsNull(TxtDecSiteIlliciteURLAdresse)
                  Case True

                       MsgBox "Le champ adresse URL site illicite est obligatoire", , Me.Caption

                       TxtDecSiteIlliciteURLAdresse.SetFocus

                       Cancel = True

                       GoTo Exit_Form_BeforeUpdate

                  Case False

               End Select

            Case 2, 1

                 Select Case IsNull(TxtDecFraudeDate)
                    Case True

                       MsgBox "Le champ date d'opération fraude est obligatoire", , Me.Caption

                       TxtDecFraudeDate.SetFocus

                       Cancel = True

                       GoTo Exit_Form_BeforeUpdate

                    Case False

                 End Select

                 Select Case IsNull(CmbDecFraudeTypOpeNumero)
                    Case True

                       MsgBox "Le champ type opération fraude est obligatoire", , Me.Caption

                       CmbDecFraudeTypOpeNumero.SetFocus

                       Cancel = True

                       GoTo Exit_Form_BeforeUpdate

                    Case False

                 End Select

                 Select Case IsNull(TxtDecFraudeMontant)
                    Case True

                       MsgBox "Le champ montant fraude est obligatoire", , Me.Caption

                       TxtDecFraudeMontant.SetFocus

                       Cancel = True

                       GoTo Exit_Form_BeforeUpdate

                    Case False

                 End Select

        End Select

      Case Else

         Cancel = True

   End Select

Exit_Form_BeforeUpdate:

   Exit Sub

Err_Form_BeforeUpdate:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_Form_BeforeUpdate
End Sub

Private Sub OgtDeclaration_Change()
   Dim PgePage As Page

   On Error GoTo Err_OgtDeclaration_Change

   For Each PgePage In OgtDeclaration.Pages

      Select Case OgtDeclaration
         Case PgePage.PageIndex

            PgePage.Picture = CurrentProject.Path & "\Fleche.BMP"

         Case Else

            PgePage.Picture = vbNullString

      End Select

   Next

Exit_OgtDeclaration_Change:

   Exit Sub

Err_OgtDeclaration_Change:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_OgtDeclaration_Change
End Sub

Private Sub TxtDecDate_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_TxtDecDate_KeyDown

   Select Case KeyCode
      Case 9

         Select Case IsNumeric(Right(TxtDecDate.Text, 4)) Or Not IsNumeric(Left(TxtDecDate.Text, 2))
            Case True

            Case False

               TxtDecDate.Text = Left(TxtDecDate.Text, 6) & Year(Date)

         End Select

      Case Else

   End Select

Exit_TxtDecDate_KeyDown:

   Exit Sub

Err_TxtDecDate_KeyDown:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtDecDate_KeyDown
End Sub

Private Sub TxtDecVictimeNaissanceDate_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_TxtDecVictimeNaissanceDate_KeyDown

   Select Case KeyCode
      Case 9

         Select Case IsNumeric(Right(TxtDecVictimeNaissanceDate.Text, 4)) Or Not IsNumeric(Left(TxtDecVictimeNaissanceDate.Text, 2))
            Case True

            Case False

               TxtDecVictimeNaissanceDate.Text = Left(TxtDecVictimeNaissanceDate.Text, 6) & Year(Date)

         End Select

      Case Else

   End Select

Exit_TxtDecVictimeNaissanceDate_KeyDown:

   Exit Sub

Err_TxtDecVictimeNaissanceDate_KeyDown:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtDecVictimeNaissanceDate_KeyDown
End Sub

Private Sub TxtDecBeneficiaireNaissanceDate_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_TxtDecBeneficiaireNaissanceDate_KeyDown

   Select Case KeyCode
      Case 9

         Select Case IsNumeric(Right(TxtDecBeneficiaireNaissanceDate.Text, 4)) Or Not IsNumeric(Left(TxtDecBeneficiaireNaissanceDate.Text, 2))
            Case True

            Case False

               TxtDecBeneficiaireNaissanceDate.Text = Left(TxtDecBeneficiaireNaissanceDate.Text, 6) & Year(Date)

         End Select

      Case Else

   End Select

Exit_TxtDecBeneficiaireNaissanceDate_KeyDown:

   Exit Sub

Err_TxtDecBeneficiaireNaissanceDate_KeyDown:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtDecBeneficiaireNaissanceDate_KeyDown
End Sub

Private Sub TxtDecFraudeDate_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_TxtDecFraudeDate_KeyDown

   Select Case KeyCode
      Case 9

         Select Case IsNumeric(Right(TxtDecFraudeDate.Text, 4)) Or Not IsNumeric(Left(TxtDecFraudeDate.Text, 2))
            Case True

            Case False

               TxtDecFraudeDate.Text = Left(TxtDecFraudeDate.Text, 6) & Year(Date)

         End Select

      Case Else

   End Select

Exit_TxtDecFraudeDate_KeyDown:

   Exit Sub

Err_TxtDecFraudeDate_KeyDown:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtDecFraudeDate_KeyDown
End Sub

Private Sub TxtDecFraudeDecaissementDate_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_TxtDecFraudeDecaissementDate_KeyDown

   Select Case KeyCode
      Case 9

         Select Case IsNumeric(Right(TxtDecFraudeDecaissementDate.Text, 4)) Or Not IsNumeric(Left(TxtDecFraudeDecaissementDate.Text, 2))
            Case True

            Case False

               TxtDecFraudeDecaissementDate.Text = Left(TxtDecFraudeDecaissementDate.Text, 6) & Year(Date)

         End Select

      Case Else

   End Select

Exit_TxtDecFraudeDecaissementDate_KeyDown:

   Exit Sub

Err_TxtDecFraudeDecaissementDate_KeyDown:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtDecFraudeDecaissementDate_KeyDown
End Sub

Private Sub TxtDecDepotPlainteClientDate_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_TxtDecDepotPlainteClientDate_KeyDown

   Select Case KeyCode
      Case 9

         Select Case IsNumeric(Right(TxtDecDepotPlainteClientDate.Text, 4)) Or Not IsNumeric(Left(TxtDecDepotPlainteClientDate.Text, 2))
            Case True

            Case False

               TxtDecDepotPlainteClientDate.Text = Left(TxtDecDepotPlainteClientDate.Text, 6) & Year(Date)

         End Select

      Case Else

   End Select

Exit_TxtDecDepotPlainteClientDate_KeyDown:

   Exit Sub

Err_TxtDecDepotPlainteClientDate_KeyDown:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtDecDepotPlainteClientDate_KeyDown
End Sub

Private Sub TxtDecDepotPlainteEntiteDate_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_TxtDecDepotPlainteEntiteDate_KeyDown

   Select Case KeyCode
      Case 9

         Select Case IsNumeric(Right(TxtDecDepotPlainteEntiteDate.Text, 4)) Or Not IsNumeric(Left(TxtDecDepotPlainteEntiteDate.Text, 2))
            Case True

            Case False

               TxtDecDepotPlainteEntiteDate.Text = Left(TxtDecDepotPlainteEntiteDate.Text, 6) & Year(Date)

         End Select

      Case Else

   End Select

Exit_TxtDecDepotPlainteEntiteDate_KeyDown:

   Exit Sub

Err_TxtDecDepotPlainteEntiteDate_KeyDown:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtDecDepotPlainteEntiteDate_KeyDown
End Sub

Private Sub TxtDecSiteIlliciteDate_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_TxtDecSiteIlliciteDate_KeyDown

   Select Case KeyCode
      Case 9

         Select Case IsNumeric(Right(TxtDecSiteIlliciteDate.Text, 4)) Or Not IsNumeric(Left(TxtDecSiteIlliciteDate.Text, 2))
            Case True

            Case False

               TxtDecSiteIlliciteDate.Text = Left(TxtDecSiteIlliciteDate.Text, 6) & Year(Date)

         End Select

      Case Else

   End Select

Exit_TxtDecSiteIlliciteDate_KeyDown:

   Exit Sub

Err_TxtDecSiteIlliciteDate_KeyDown:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtDecSiteIlliciteDate_KeyDown
End Sub

Private Sub TxtDecLiaisonSFDate_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_TxtDecLiaisonSFDate_KeyDown

   Select Case KeyCode
      Case 9

         Select Case IsNumeric(Right(TxtDecLiaisonSFDate.Text, 4)) Or Not IsNumeric(Left(TxtDecLiaisonSFDate.Text, 2))
            Case True

            Case False

               TxtDecLiaisonSFDate.Text = Left(TxtDecLiaisonSFDate.Text, 6) & Year(Date)

         End Select

      Case Else

   End Select

Exit_TxtDecLiaisonSFDate_KeyDown:

   Exit Sub

Err_TxtDecLiaisonSFDate_KeyDown:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_TxtDecLiaisonSFDate_KeyDown
End Sub

Private Sub CmbDecTypDecNumero_NotInList(NewData As String, Response As Integer)
   Dim RsTypeDeclaration As DAO.Recordset

   On Error GoTo Err_CmbDecTypDecNumero_NotInList

   Select Case MsgBox("Est-ce un nouveaux type de déclaration ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsTypeDeclaration = CurrentDb.OpenRecordset("TBLTYPEDECLARATION")

         RsTypeDeclaration.AddNew

         RsTypeDeclaration("TypDecNom") = NewData

         RsTypeDeclaration.Update

         RsTypeDeclaration.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecTypDecNumero_NotInList:

   Set RsTypeDeclaration = Nothing

   Exit Sub

Err_CmbDecTypDecNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecTypDecNumero_NotInList
End Sub

Private Sub CmbDecDeclarantEntNumero_NotInList(NewData As String, Response As Integer)
   Dim RsEntite As DAO.Recordset

   On Error GoTo Err_CmbDecDeclarantEntNumero_NotInList

   Select Case MsgBox("Est-ce une nouvelle entité ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsEntite = CurrentDb.OpenRecordset("TBLENTITE")

         RsEntite.AddNew

         RsEntite("EntNom") = NewData

         RsEntite.Update

         RsEntite.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecDeclarantEntNumero_NotInList:

   Set RsEntite = Nothing

   Exit Sub

Err_CmbDecDeclarantEntNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecDeclarantEntNumero_NotInList
End Sub

Private Sub CmbDecVictimeTypCptNumero_NotInList(NewData As String, Response As Integer)
   Dim RsTypeCompte As DAO.Recordset

   On Error GoTo Err_CmbDecVictimeTypCptNumero_NotInList

   Select Case MsgBox("Est-ce un nouveaux type de compte ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsTypeCompte = CurrentDb.OpenRecordset("TBLTYPECOMPTE")

         RsTypeCompte.AddNew

         RsTypeCompte("TypCptNom") = NewData

         RsTypeCompte.Update

         RsTypeCompte.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecVictimeTypCptNumero_NotInList:

   Set RsTypeCompte = Nothing

   Exit Sub

Err_CmbDecVictimeTypCptNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecVictimeTypCptNumero_NotInList
End Sub

Private Sub CmbDecVictimeMarNumero_NotInList(NewData As String, Response As Integer)
   Dim RsMarche As DAO.Recordset

   On Error GoTo Err_CmbDecVictimeMarNumero_NotInList

   Select Case MsgBox("Est-ce un nouveaux marché ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsMarche = CurrentDb.OpenRecordset("TBLMARCHE")

         RsMarche.AddNew

         RsMarche("MarNom") = NewData

         RsMarche.Update

         RsMarche.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecVictimeMarNumero_NotInList:

   Set RsMarche = Nothing

   Exit Sub

Err_CmbDecVictimeMarNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecVictimeMarNumero_NotInList
End Sub

Private Sub CmbDecBeneficiaireTypBenNumero_NotInList(NewData As String, Response As Integer)
   Dim RsTypeBeneficiaire As DAO.Recordset

   On Error GoTo Err_CmbDecBeneficiaireTypBenNumero_NotInList

   Select Case MsgBox("Est-ce un nouveaux type de bénéficiaire ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsTypeBeneficiaire = CurrentDb.OpenRecordset("TBLTYPEBENEFICIAIRE")

         RsTypeBeneficiaire.AddNew

         RsTypeBeneficiaire("TypBenNom") = NewData

         RsTypeBeneficiaire.Update

         RsTypeBeneficiaire.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecBeneficiaireTypBenNumero_NotInList:

   Set RsTypeBeneficiaire = Nothing

   Exit Sub

Err_CmbDecBeneficiaireTypBenNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecBeneficiaireTypBenNumero_NotInList
End Sub

Private Sub CmbDecFraudeTypOpeNumero_NotInList(NewData As String, Response As Integer)
   Dim RsTypeOperation As DAO.Recordset

   On Error GoTo Err_CmbDecFraudeTypOpeNumero_NotInList

   Select Case MsgBox("Est-ce un nouveaux type d'opération ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsTypeOperation = CurrentDb.OpenRecordset("TBLTYPEOPERATION")

         RsTypeOperation.AddNew

         RsTypeOperation("TypOpeNom") = NewData

         RsTypeOperation.Update

         RsTypeOperation.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecFraudeTypOpeNumero_NotInList:

   Set RsTypeOperation = Nothing

   Exit Sub

Err_CmbDecFraudeTypOpeNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecFraudeTypOpeNumero_NotInList
End Sub

Private Sub CmbDecFraudeTypFrdNumero_NotInList(NewData As String, Response As Integer)
   Dim RsTypeFraude As DAO.Recordset

   On Error GoTo Err_CmbDecFraudeTypFrdNumero_NotInList

   Select Case MsgBox("Est-ce un nouveaux type de fraude ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsTypeFraude = CurrentDb.OpenRecordset("TBLTYPEFRAUDE")

         RsTypeFraude.AddNew

         RsTypeFraude("TypFrdNom") = NewData

         RsTypeFraude.Update

         RsTypeFraude.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecFraudeTypFrdNumero_NotInList:

   Set RsTypeFraude = Nothing

   Exit Sub

Err_CmbDecFraudeTypFrdNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecFraudeTypFrdNumero_NotInList
End Sub

Private Sub CmbDecFraudeSouIPNumero_NotInList(NewData As String, Response As Integer)
   Dim RsSourceIP As DAO.Recordset

   On Error GoTo Err_CmbDecFraudeSouIPNumero_NotInList

   Select Case MsgBox("Est-ce une nouvelle source d'IP ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsSourceIP = CurrentDb.OpenRecordset("TBLSOURCEIP")

         RsSourceIP.AddNew

         RsSourceIP("SouIPNom") = NewData

         RsSourceIP.Update

         RsSourceIP.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecFraudeSouIPNumero_NotInList:

   Set RsSourceIP = Nothing

   Exit Sub

Err_CmbDecFraudeSouIPNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecFraudeSouIPNumero_NotInList
End Sub

Private Sub CmbDecDepotPlainteClientAutNumero_NotInList(NewData As String, Response As Integer)
   Dim RsAutorite As DAO.Recordset

   On Error GoTo Err_CmbDecDepotPlainteClientAutNumero_NotInList

   Select Case MsgBox("Est-ce une nouvelle autorité ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsAutorite = CurrentDb.OpenRecordset("TBLAUTORITE")

         RsAutorite.AddNew

         RsAutorite("AutNom") = NewData

         RsAutorite.Update

         RsAutorite.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecDepotPlainteClientAutNumero_NotInList:

   Set RsAutorite = Nothing

   Exit Sub

Err_CmbDecDepotPlainteClientAutNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecDepotPlainteClientAutNumero_NotInList
End Sub

Private Sub CmbDecDepotPlainteEntiteAutNumero_NotInList(NewData As String, Response As Integer)
   Dim RsAutorite As DAO.Recordset

   On Error GoTo Err_CmbDecDepotPlainteEntiteAutNumero_NotInList

   Select Case MsgBox("Est-ce une nouvelle autorité ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsAutorite = CurrentDb.OpenRecordset("TBLAUTORITE")

         RsAutorite.AddNew

         RsAutorite("AutNom") = NewData

         RsAutorite.Update

         RsAutorite.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecDepotPlainteEntiteAutNumero_NotInList:

   Set RsAutorite = Nothing

   Exit Sub

Err_CmbDecDepotPlainteEntiteAutNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecDepotPlainteEntiteAutNumero_NotInList
End Sub

Private Sub CmbDecSiteIlliciteStaSitNumero_NotInList(NewData As String, Response As Integer)
   Dim RsStatutSite As DAO.Recordset

   On Error GoTo Err_CmbDecSiteIlliciteStaSitNumero_NotInList

   Select Case MsgBox("Est-ce un nouveaux statut ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsStatutSite = CurrentDb.OpenRecordset("TBLSTATUTSITE")

         RsStatutSite.AddNew

         RsStatutSite("StaSitNom") = NewData

         RsStatutSite.Update

         RsStatutSite.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecSiteIlliciteStaSitNumero_NotInList:

   Set RsStatutSite = Nothing

   Exit Sub

Err_CmbDecSiteIlliciteStaSitNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecSiteIlliciteStaSitNumero_NotInList
End Sub

Private Sub CmbDecSiteIlliciteSouSitNumero_NotInList(NewData As String, Response As Integer)
   Dim RsSourceSite As DAO.Recordset

   On Error GoTo Err_CmbDecSiteIlliciteSouSitNumero_NotInList

   Select Case MsgBox("Est-ce une nouvelle source de détection ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsSourceSite = CurrentDb.OpenRecordset("TBLSOURCESITE")

         RsSourceSite.AddNew

         RsSourceSite("SouSitNom") = NewData

         RsSourceSite.Update

         RsSourceSite.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecSiteIlliciteSouSitNumero_NotInList:

   Set RsSourceSite = Nothing

   Exit Sub

Err_CmbDecSiteIlliciteSouSitNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecSiteIlliciteSouSitNumero_NotInList
End Sub

Private Sub CmbDecSiteIlliciteLocSitNumero_NotInList(NewData As String, Response As Integer)
   Dim RsLocalisationSite As DAO.Recordset

   On Error GoTo Err_CmbDecSiteIlliciteLocSitNumero_NotInList

   Select Case MsgBox("Est-ce une nouvelle localisation de site ?", vbYesNo, Me.Caption)
      Case vbYes

         Set RsLocalisationSite = CurrentDb.OpenRecordset("TBLLOCALISATIONSITE")

         RsLocalisationSite.AddNew

         RsLocalisationSite("LocSitNom") = NewData

         RsLocalisationSite.Update

         RsLocalisationSite.Close

         Response = acDataErrAdded

      Case vbNo

         Response = acDataErrContinue

   End Select

Exit_CmbDecSiteIlliciteLocSitNumero_NotInList:

   Set RsLocalisationSite = Nothing

   Exit Sub

Err_CmbDecSiteIlliciteLocSitNumero_NotInList:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmbDecSiteIlliciteLocSitNumero_NotInList
End Sub

Public Sub CmdAjouter_Click()

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.Recordset.RecordCount
      Case 0

      Case Else

         TxtDecDate.SetFocus

   End Select

   ChangeMode Me, MODEAJOUTER

   DoCmd.GoToRecord , , acNewRec

Exit_CmdAjouter_Click:

   Exit Sub

Err_CmdAjouter_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdAjouter_Click
End Sub

Public Sub CmdAnnuler_Click()

   On Error GoTo Err_CmdAnnuler_Click

   TxtDecDate.SetFocus

   ChangeMode Me, MODEANNULER

   On Error Resume Next

   DoCmd.RunCommand acCmdUndo

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdAnnuler_Click

   End Select

   ChangeMode Me, MODEDEPLACER

Exit_CmdAnnuler_Click:

   Exit Sub

Err_CmdAnnuler_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdAnnuler_Click
End Sub

Public Sub CmdModifier_Click()

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.Recordset.RecordCount
      Case 0

      Case Else

         TxtDecDate.SetFocus

         ChangeMode Me, MODEMODIFIER

         DecDateModification = Date

   End Select

Exit_CmdModifier_Click:

   Exit Sub

Err_CmdModifier_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdModifier_Click
End Sub

Public Sub CmdSupprimer_Click()

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.Recordset.RecordCount
      Case 0

      Case Else

         TxtDecDate.SetFocus

         ChangeMode Me, MODESUPPRIMER

         On Error Resume Next

         DoCmd.RunCommand acCmdDeleteRecord

         Select Case Err.Number
            Case 0

            Case 2046, 2501

               Err.Clear

            Case Else

               GoTo Err_CmdSupprimer_Click

         End Select

         ChangeMode Me, MODEDEPLACER

   End Select

Exit_CmdSupprimer_Click:

   Exit Sub

Err_CmdSupprimer_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdSupprimer_Click
End Sub

Public Sub CmdValider_Click()
  Dim CtlControl As Control

   On Error GoTo Err_CmdValider_Click

   TxtDecDate.SetFocus

   ChangeMode Me, MODEVALIDER

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

         Select Case IntModeOld
            Case MODEAJOUTER

               CreeMailDeclaration "sca.mail@credit-agricole-sa.fr", TxtDecDeclarantMail, "Déclaration De Fraude", DLookup("ParValeur", "TBLPARAMETRE", "ParCode='TXTMAI'"), "DecNumero=" & Me.DecNumero

            Case Else

         End Select

         ChangeMode Me, MODEDEPLACER

         Select Case IntModeGeneral
            Case MODEDEPLACER

            Case MODEAJOUTER

               CmdAjouter_Click

            Case MODEMODIFIER

               CmdQuitter_Click

         End Select

      Case 2046, 2501, 3021

         Err.Clear

         ChangeMode Me, IntModeOld

      Case Else

         GoTo Err_CmdValider_Click

   End Select

Exit_CmdValider_Click:

   Set CtlControl = Nothing

   Exit Sub

Err_CmdValider_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdValider_Click
End Sub

Public Sub CmdQuitter_Click()

   On Error GoTo Err_CmdQuitter_Click

   Select Case IntModeGeneral
      Case MODEDEPLACER

      Case MODEAJOUTER, MODEMODIFIER

               CmdAnnuler_Click

   End Select

   Select Case IntMode
      Case MODEDEPLACER

      Case MODEAJOUTER, MODEMODIFIER

         CmdAnnuler_Click

   End Select

   DoCmd.RunCommand acCmdClose

Exit_CmdQuitter_Click:

   Exit Sub

Err_CmdQuitter_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdQuitter_Click
End Sub

Public Sub CmdImprimer_Click()

   On Error GoTo Err_CmdImprimer_Click

   DoCmd.OpenReport "EtaDeclaration", acViewPreview, , "DecNumero=" & Me.DecNumero, acWindowNormal

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdImprimer_Click
End Sub
