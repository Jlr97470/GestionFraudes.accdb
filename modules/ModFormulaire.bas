Attribute VB_Name = "ModFormulaire"
Option Explicit

Public Const MODEDEPLACER = 0
Public Const MODEAJOUTER = 1
Public Const MODEMODIFIER = 2
Public Const MODESUPPRIMER = 3
Public Const MODEVALIDER = 4
Public Const MODEANNULER = 5

Public Function ChangeMode(ByRef FrmForm As Form, ByVal IntMode As Integer) As Boolean

   On Error GoTo Err_ChangeMode

   ChangeMode = True

   FrmForm.IntModeOld = FrmForm.IntMode

   FrmForm.IntMode = IntMode

   Select Case IntMode
      Case MODEDEPLACER

       '  FrmForm.BteControl.BackColor = vbBlack

       '  FrmForm.BteControl.BackStyle = 0

         FrmForm.AllowAdditions = False

         FrmForm.AllowEdits = False

         FrmForm.AllowDeletions = False

         FrmForm.NavigationButtons = True

         FrmForm.CmdAjouter.Enabled = True

         FrmForm.CmdModifier.Enabled = True

         FrmForm.CmdSupprimer.Enabled = True

         FrmForm.CmdValider.Enabled = False

         FrmForm.CmdAnnuler.Enabled = False

         FrmForm.CmdImprimer.Enabled = True

         FrmForm.CmdQuitter.Enabled = True

      Case MODEAJOUTER

     '    FrmForm.BteControl.BackColor = vbGreen

      '   FrmForm.BteControl.BackStyle = 1

         FrmForm.AllowAdditions = True

         FrmForm.AllowEdits = False

         FrmForm.AllowDeletions = False

         FrmForm.NavigationButtons = False

         FrmForm.CmdValider.Enabled = True

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdAnnuler.Enabled = True

         FrmForm.CmdImprimer.Enabled = False

         FrmForm.CmdQuitter.Enabled = True

      Case MODEMODIFIER

      '   FrmForm.BteControl.BackColor = vbBlue

     '    FrmForm.BteControl.BackStyle = 1

         FrmForm.AllowAdditions = False

         FrmForm.AllowEdits = True

         FrmForm.AllowDeletions = False

         FrmForm.NavigationButtons = False

         FrmForm.CmdValider.Enabled = True

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdAnnuler.Enabled = True

         FrmForm.CmdImprimer.Enabled = False

         FrmForm.CmdQuitter.Enabled = True

      Case MODESUPPRIMER

    '     FrmForm.BteControl.BackColor = vbRed

     '    FrmForm.BteControl.BackStyle = 1

         FrmForm.AllowAdditions = False

         FrmForm.AllowEdits = False

         FrmForm.AllowDeletions = True

         FrmForm.NavigationButtons = False

         FrmForm.CmdSupprimer.Enabled = True

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdValider.Enabled = False

         FrmForm.CmdAnnuler.Enabled = False

         FrmForm.CmdImprimer.Enabled = False

         FrmForm.CmdQuitter.Enabled = True

      Case MODEVALIDER

    '     FrmForm.BteControl.BackColor = vbWhite

    '     FrmForm.BteControl.BackStyle = 1

         FrmForm.NavigationButtons = False

         FrmForm.CmdValider.Enabled = True

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdAnnuler.Enabled = True

         FrmForm.CmdImprimer.Enabled = False

         FrmForm.CmdQuitter.Enabled = True

      Case MODEANNULER

   '      FrmForm.BteControl.BackColor = vbMagenta

     '    FrmForm.BteControl.BackStyle = 1

         FrmForm.NavigationButtons = False

         FrmForm.CmdAnnuler.Enabled = True

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdValider.Enabled = False

         FrmForm.CmdImprimer.Enabled = False

         FrmForm.CmdQuitter.Enabled = True

   End Select

Exit_ChangeMode:

   Exit Function

Err_ChangeMode:

   ChangeMode = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_ChangeMode
End Function

Public Function OuvreFrmDeclaration(ByVal IntMode As Integer, ByVal StrCaption As String, Optional ByVal StrFiltre As String) As Boolean

   On Error GoTo Err_OuvreFrmDeclaration

   OuvreFrmDeclaration = True

   DoCmd.OpenForm "FrmDeclaration", acNormal, , StrFiltre, acFormEdit, acWindowNormal

   Forms("FrmDeclaration").Caption = StrCaption

   Forms("FrmDeclaration").IntModeGeneral = IntMode

   Forms("FrmDeclaration").CmdAjouter.Visible = False

   Forms("FrmDeclaration").CmdModifier.Visible = False

   Forms("FrmDeclaration").CmdSupprimer.Visible = False

   Forms("FrmDeclaration").CmdAnnuler.Visible = False

   Select Case IntMode
      Case MODEAJOUTER

         Forms("FrmDeclaration").CmdAjouter_Click

      Case MODEMODIFIER

         Forms("FrmDeclaration").CmdModifier_Click

   End Select

Exit_OuvreFrmDeclaration:

   Exit Function

Err_OuvreFrmDeclaration:

   OuvreFrmDeclaration = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_OuvreFrmDeclaration
End Function

Public Function OuvreFrmDeclarationRecherche(ByVal StrCaption As String) As Boolean

   On Error GoTo Err_OuvreFrmDeclarationRecherche

   OuvreFrmDeclarationRecherche = True

   DoCmd.OpenForm "FrmDeclarationRecherche", acNormal, , , acFormEdit, acWindowNormal

   Forms("FrmDeclarationRecherche").Caption = StrCaption

Exit_OuvreFrmDeclarationRecherche:

   Exit Function

Err_OuvreFrmDeclarationRecherche:

   OuvreFrmDeclarationRecherche = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_OuvreFrmDeclarationRecherche
End Function

Public Function OuvreFrmDeclarationAccuse(ByVal StrCaption As String) As Boolean

   On Error GoTo Err_OuvreFrmDeclarationAccuse

   OuvreFrmDeclarationAccuse = True

   DoCmd.OpenForm "FrmDeclarationAccuse", acNormal, , , acFormEdit, acWindowNormal

   Forms("FrmDeclarationAccuse").Caption = StrCaption

Exit_OuvreFrmDeclarationAccuse:

   Exit Function

Err_OuvreFrmDeclarationAccuse:

   OuvreFrmDeclarationAccuse = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_OuvreFrmDeclarationAccuse
End Function

Public Function OuvreFrmMenuListe(ByVal StrCaption As String) As Boolean

   On Error GoTo Err_OuvreFrmMenuListe

   OuvreFrmMenuListe = True

   DoCmd.OpenForm "FrmMenuListe", acNormal, , , acFormEdit, acWindowNormal

   Forms("FrmMenuListe").Caption = StrCaption

Exit_OuvreFrmMenuListe:

   Exit Function

Err_OuvreFrmMenuListe:

   OuvreFrmMenuListe = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_OuvreFrmMenuListe
End Function

Public Function FermeEtatActif() As Boolean

   On Error GoTo Err_FermeEtatActif

   DoCmd.Close acReport, Screen.ActiveReport.Name, acSavePrompt

Exit_FermeEtatActif:

   Exit Function

Err_FermeEtatActif:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_FermeEtatActif
End Function
