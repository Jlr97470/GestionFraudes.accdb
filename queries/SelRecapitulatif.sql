PARAMETERS [Formulaires]![FrmTableauBordStatistique]![TxtDateDebut] DateTime,
[Formulaires]![FrmTableauBordStatistique]![TxtDateFin] DateTime;
SELECT
  Null AS Mois,
  Sum(
    (
      (
        IIf(
          [TypOpeNumero] = 1
          And [TypDecNumero] = 1,
          1, 0
        )
      )
    )
  ) AS NbDébit,
  Sum(
    (
      (
        IIf(
          [TypOpeNumero] = 2
          And [TypDecNumero] = 1,
          1, 0
        )
      )
    )
  ) AS NbCrédiit,
  Sum(
    (
      (
        IIf([TypDecNumero] = 1, 1, 0)
      )
    )
  ) AS NbFraude,
  Sum(
    (
      (
        IIf(
          [TypOpeNumero] = 1
          And [TypDecNumero] = 1,
          [DecFraudeMontant], 0
        )
      )
    )
  ) AS SomDébit,
  Sum(
    (
      (
        IIf(
          [TypOpeNumero] = 2
          And [TypDecNumero] = 1,
          [DecFraudeMontant], 0
        )
      )
    )
  ) AS SomCrédiit,
  Sum(
    (
      (
        IIf(
          [TypOpeNumero] = 1
          And [TypDecNumero] = 1,
          [DecFraudeMontant], 0
        )
      )
    )
  ) AS MoyDébit,
  Sum(
    (
      (
        IIf(
          [TypOpeNumero] = 2
          And [TypDecNumero] = 1,
          [DecFraudeMontant], 0
        )
      )
    )
  ) AS MoyCrédiit,
  Sum(
    (
      (
        IIf(
          [TypDecNumero] = 1, [DecFraudeMontant],
          0
        )
      )
    )
  ) AS PreNet,
  Sum(
    (
      (
        IIf(
          [TypFrdNumero] In (2, 3, 4, 5, 7, 8, 10),
          1,
          0
        )
      )
    )
  ) AS TenFraudes,
  Sum(
    (
      (
        IIf(
          [TypFrdNumero] In (2, 3, 4, 5, 7, 8, 10),
          [DecFraudeMontant],
          0
        )
      )
    )
  ) AS MonFraudes,
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
  Null;
