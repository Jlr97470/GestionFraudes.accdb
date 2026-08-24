SELECT
  TBLENTITE.EntNom,
  (
    IIf(
      [TypOpeNumero] = 1
      And [TypDecNumero] = 1,
      1, 0
    )
  ) AS NbDébit,
  (
    IIf(
      [TypOpeNumero] = 1
      And [TypDecNumero] = 1,
      [DecFraudeMontant], 0
    )
  ) AS SomDébit,
  (
    IIf(
      [TypOpeNumero] = 2
      And [TypDecNumero] = 1,
      1, 0
    )
  ) AS NbCrédiit,
  (
    IIf(
      [TypOpeNumero] = 2
      And [TypDecNumero] = 1,
      [DecFraudeMontant], 0
    )
  ) AS SomCrédiit,
  (
    IIf(
      [TypDecNumero] = 1, [DecFraudeMontant],
      0
    )
  ) AS PreNet,
  (
    IIf(
      [TypFrdNumero] In (2, 3, 4, 5, 7, 8, 10),
      1,
      0
    )
  ) AS TenFraudes,
  (
    IIf(
      [TypFrdNumero] In (2, 3, 4, 5, 7, 8, 10),
      [DecFraudeMontant],
      0
    )
  ) AS MonFraudes,
  TBLDECLARATION.DecFraudeMontant AS RisGlo,
  (
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
  ) ON TBLTYPEOPERATION.TypOpeNumero = TBLDECLARATION.DecFraudeTypOpeNumero;
