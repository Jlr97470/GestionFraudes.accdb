PARAMETERS [Formulaires!FrmTableauBordStatistique!TxtDateDebut] DateTime,
[Formulaires!FrmTableauBordStatistique!TxtDateFin] DateTime;
SELECT
  CStr(
    CStr(
      Month([DecDateCreation])& " " & Year([DecDateCreation])
    )
  ) AS FraudeMois,
  (
    IIf(
      [TypOpeNumero] = 1
      And [TypDecNumero] = 1,
      [DecFraudeMontant], 0
    )
  ) AS Débit,
  (
    IIf(
      [TypOpeNumero] = 2
      And [TypDecNumero] = 1,
      [DecFraudeMontant], 0
    )
  ) AS Crédiit,
  TBLENTITE.EntNom,
  TBLTYPEFRAUDE.TypFrdNom,
  TBLDECLARATION.DecDepotPlainteClientDate,
  TBLDECLARATION.DecDepotPlainteClientLieu
FROM
  TBLTYPEOPERATION
  INNER JOIN (
    TBLTYPEDECLARATION
    INNER JOIN (
      TBLTYPEFRAUDE
      INNER JOIN (
        TBLDECLARATION
        INNER JOIN TBLENTITE ON TBLDECLARATION.DecDeclarantEntNumero = TBLENTITE.EntNumero
      ) ON TBLTYPEFRAUDE.TypFrdNumero = TBLDECLARATION.DecFraudeTypFrdNumero
    ) ON TBLTYPEDECLARATION.TypDecNumero = TBLDECLARATION.DecTypDecNumero
  ) ON TBLTYPEOPERATION.TypOpeNumero = TBLDECLARATION.DecFraudeTypOpeNumero
WHERE
  (
    (
      (TBLDECLARATION.DecFraudeDate) >= [Formulaires]![FrmTableauBordStatistique]![TxtDateDebut]
      And (TBLDECLARATION.DecFraudeDate) <= [Formulaires]![FrmTableauBordStatistique]![TxtDateFin]
    )
  );
