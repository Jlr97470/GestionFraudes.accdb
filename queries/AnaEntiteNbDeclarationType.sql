TRANSFORM
  First(
    DCount(
      "DecNumero", "TBLDECLARATION", "DecTypDecNumero=" & [TypDecNumero] & " AND DecDeclarantEntNumero=" & [EntNumero]
    )
  ) AS NbDeclaration
SELECT
  SelEntiteTypeDeclaration.EntNom
FROM
  SelEntiteTypeDeclaration
GROUP BY
  SelEntiteTypeDeclaration.EntNom
PIVOT
  SelEntiteTypeDeclaration.TypDecNom;
