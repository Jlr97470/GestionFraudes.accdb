TRANSFORM
  First(
    DCount(
      "DecNumero", "TBLDECLARATION", "DecTypDecNumero=" & [TypDecNumero]
    )
  ) AS NbDeclaration
SELECT
  "TOTAL" AS [TOTAL GENERAL],
  DCount("DecNumero", "TBLDECLARATION") AS TOTAL
FROM
  TBLTYPEDECLARATION
GROUP BY
  "TOTAL",
  DCount("DecNumero", "TBLDECLARATION")
PIVOT
  TBLTYPEDECLARATION.TypDecNom;
