INSERT INTO TBLDECLARATION (
  DecDate, DecTypDecNumero, DecDeclarantEntNumero,
  DecDeclarantNom, DecDeclarantPrenom,
  DecDeclarantQualite, DecDeclarantTelephone1,
  DecDeclarantMail, DecFraudeDate,
  DecFraudeMontant, DecFraudeTypOpeNumero,
  DecFraudeTypFrdNumero, DecDepotPlainteClient,
  DecDepotPlainteClientDate, DecDepotPlainteClientLieu
)
SELECT
  Date() AS [Date],
  1 AS TypDec,
  TBLENTITE.EntNumero,
  "INCONNUE" AS Nom,
  "INCONNUE" AS Prenom,
  "INCONNUE" AS Qualite,
  "9999999999" AS Telephone,
  "INCONNUE" AS Mail,
  LSTDEC.[Date Fraude],
  IIf(
    Not IsNull([Débit]),
    [Débit],
    [Crédit]
  ) AS Montant,
  IIf(
    Not IsNull([Débit]),
    1,
    2
  ) AS TypOpe,
  TBLTYPEFRAUDE.TypFrdNumero,
  IIf(
    IsNull([Date Dépôt]),
    False,
    True
  ) AS Client,
  IIf(
    IsDate([Date Dépôt]),
    [Date Dépôt],
    Null
  ) AS DateDepot,
  LSTDEC.[Lieu Dépôt]
FROM
  (
    LSTDEC
    INNER JOIN TBLENTITE ON LSTDEC.Entite = TBLENTITE.EntNom
  )
  LEFT JOIN TBLTYPEFRAUDE ON LSTDEC.[Type Fraude] = TBLTYPEFRAUDE.TypFrdNom;
