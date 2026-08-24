PARAMETERS [Formulaires]![FrmTableauBordStatistique]![TxtDateDebut] DateTime,
[Formulaires]![FrmTableauBordStatistique]![TxtDateFin] DateTime;
SELECT
  CStr(
    CStr(
      MonthName(
        Month([DecDateCreation])
      )& " " & Year([DecDateCreation])
    )
  ) AS FraudeMois,
  Sum(
    (
      IIf(
        [TypOpeNumero] = 1
        And [TypDecNumero] = 1,
        1, 0
      )
    )
  ) AS NbDebitAverees,
  Sum(
    (
      IIf(
        [TypOpeNumero] = 2
        And [TypDecNumero] = 1,
        1, 0
      )
    )
  ) AS NbCreditAverees,
  Sum(
    (
      IIf([TypDecNumero] = 1, 1, 0)
    )
  ) AS NbAverees,
  Sum(
    (
      IIf(
        [TypOpeNumero] = 1
        And [TypDecNumero] = 1,
        [DecFraudeMontant], 0
      )
    )
  ) AS MontantDebitAverees,
  Sum(
    (
      IIf(
        [TypOpeNumero] = 2
        And [TypDecNumero] = 1,
        [DecFraudeMontant], 0
      )
    )
  ) AS MontantCreditAverees,
  Avg(
    (
      IIf(
        [TypOpeNumero] = 1
        And [TypDecNumero] = 1,
        [DecFraudeMontant], 0
      )
    )
  ) AS MoyenneDebitAverees,
  Avg(
    (
      IIf(
        [TypOpeNumero] = 2
        And [TypDecNumero] = 1,
        [DecFraudeMontant], 0
      )
    )
  ) AS MoyenneCreditAverees,
  (
    SELECT
      COUNT(*)
    FROM
      (
        SELECT DISTINCT
          DecDeclarantEntNumero
        FROM
          TBLDECLARATION
        WHERE
          DecFraudeDate >= [Formulaires]![FrmTableauBordStatistique]![TxtDateDebut]
          AND DecFraudeDate <= [Formulaires]![FrmTableauBordStatistique]![TxtDateFin]
          AND [DecTypDecNumero] = 1
      )
  ) AS NbEntiteAverees,
  Sum(
    (
      IIf(
        [TypDecNumero] = 1, [DecFraudeMontant],
        0
      )
    )
  ) AS PreNet,
  Sum(
    (
      IIf(
        [TypFrdNumero] In (2, 3, 4, 5, 7, 8, 10),
        1,
        0
      )
    )
  ) AS TenFraudes,
  Sum(
    (
      IIf(
        [TypFrdNumero] In (2, 3, 4, 5, 7, 8, 10),
        [DecFraudeMontant],
        0
      )
    )
  ) AS MonFraudes,
  (
    SELECT
      COUNT(*)
    FROM
      (
        SELECT DISTINCT
          DecDeclarantEntNumero
        FROM
          TBLDECLARATION
        WHERE
          DecFraudeDate >= [Formulaires]![FrmTableauBordStatistique]![TxtDateDebut]
          AND DecFraudeDate <= [Formulaires]![FrmTableauBordStatistique]![TxtDateFin]
      )
  ) AS NbCrTotal,
  Sum(
    TBLDECLARATION.DecFraudeMontant
  ) AS RisGlo
FROM
  TBLTYPEOPERATION
  INNER JOIN (
    TBLTYPEDECLARATION
    INNER JOIN (
      TBLTYPEFRAUDE
      INNER JOIN TBLDECLARATION ON TBLTYPEFRAUDE.TypFrdNumero = TBLDECLARATION.DecFraudeTypFrdNumero
    ) ON TBLTYPEDECLARATION.TypDecNumero = TBLDECLARATION.DecTypDecNumero
  ) ON TBLTYPEOPERATION.TypOpeNumero = TBLDECLARATION.DecFraudeTypOpeNumero
WHERE
  (
    (
      (TBLDECLARATION.DecFraudeDate) >= [Formulaires]![FrmTableauBordStatistique]![TxtDateDebut]
      And (TBLDECLARATION.DecFraudeDate) <= [Formulaires]![FrmTableauBordStatistique]![TxtDateFin]
    )
  )
GROUP BY
  CStr(
    CStr(
      MonthName(
        Month([DecDateCreation])
      )& " " & Year([DecDateCreation])
    )
  );
