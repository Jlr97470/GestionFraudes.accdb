Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ValideMenu()

   On Error GoTo Err_ValideMenu

   DoCmd.RunMacro LstMenu.Column(3)

Exit_ValideMenu:

   Exit Sub

Err_ValideMenu:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_ValideMenu
End Sub

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   TxtMenuGroupe = CtbMenu.Value

   LstMenu.Requery

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_Form_Load
End Sub

Private Sub CtbMenu_Change()

   On Error GoTo Err_CtbMenu_Change

   TxtMenuGroupe = CtbMenu.Value

   LstMenu.Requery

Exit_CtbMenu_Change:

   Exit Sub

Err_CtbMenu_Change:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CtbMenu_Change
End Sub

Private Sub LstMenu_DblClick(Cancel As Integer)

   On Error GoTo Err_LstMenu_DblClick

   ValideMenu

Exit_LstMenu_DblClick:

   Exit Sub

Err_LstMenu_DblClick:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_LstMenu_DblClick
End Sub

Private Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   ValideMenu

Exit_CmdValider_Click:

   Exit Sub

Err_CmdValider_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdValider_Click
End Sub

Private Sub CmdQuitter_Click()
   On Error GoTo Err_CmdQuitter_Click

   Application.Quit

Exit_CmdQuitter_Click:

   Exit Sub

Err_CmdQuitter_Click:

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_CmdQuitter_Click
End Sub
