SELECT
  TBLENTITE.EntNom,
  Sum(
    IIf(
      [TypOpeNumero] = 1
      And [TypDecNumero] = 1,
      1, 0
    )
  ) AS NbDébit,
  Sum(
    IIf(
      [TypOpeNumero] = 1
      And [TypDecNumero] = 1,
      [DecFraudeMontant], 0
    )
  ) AS SomDébit,
  Sum(
    IIf(
      [TypOpeNumero] = 2
      And [TypDecNumero] = 1,
      1, 0
    )
  ) AS NbCrédiit,
  Sum(
    IIf(
      [TypOpeNumero] = 2
      And [TypDecNumero] = 1,
      [DecFraudeMontant], 0
    )
  ) AS SomCrédiit,
  Sum(
    IIf(
      [TypDecNumero] = 1, [DecFraudeMontant],
      0
    )
  ) AS PreNet,
  Sum(
    IIf([TypDecNumero] <> 1, 1, 0)
  ) AS NbTenFraudes,
  Sum(
    IIf(
      [TypDecNumero] <> 1, [DecFraudeMontant],
      0
    )
  ) AS NbTenMonFraudes,
  Sum(
    TBLDECLARATION.DecFraudeMontant
  ) AS RisGlo,
  Sum(
    IIf([TypDecNumero] = 3, 1, 0)
  ) AS SitIlli
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
GROUP BY
  TBLENTITE.EntNom;
