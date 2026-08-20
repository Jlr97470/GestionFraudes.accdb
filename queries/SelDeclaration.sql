SELECT
  TBLDECLARATION.DecNumero,
  TBLDECLARATION.DecDateCreation,
  TBLDECLARATION.DecDateModification,
  TBLDECLARATION.DecDate,
  TBLTYPEDECLARATION.TypDecNom AS DecTypDecNom,
  TBLENTITE.EntNom AS DecDeclarantEntNom,
  TBLDECLARATION.DecDeclarantNom,
  TBLDECLARATION.DecDeclarantPrenom,
  TBLDECLARATION.DecDeclarantQualite,
  TBLDECLARATION.DecDeclarantTelephone1,
  TBLDECLARATION.DecDeclarantTelephone2,
  TBLDECLARATION.DecDeclarantMail,
  TBLDECLARATION.DecVictimeNom,
  TBLDECLARATION.DecVictimePrenom,
  TBLDECLARATION.DecVictimeNaissanceDate,
  TBLDECLARATION.DecVictimeNaissanceLieu,
  TBLDECLARATION.DecVictimeAdresse,
  TBLDECLARATION.DecVictimeCodePostal,
  TBLDECLARATION.DecVictimeVille,
  TBLDECLARATION.DecVictimePays,
  TBLDECLARATION.DecVictimeTelephone1,
  TBLDECLARATION.DecVictimeTelephone2,
  TBLDECLARATION.DecVictimeRIB,
  TBLTYPECOMPTE.TypCptNom AS DecVictimeTypCptNom,
  TBLDECLARATION.DecVictimeAgence,
  TBLDECLARATION.DecVictimeAgenceAdresse,
  TBLDECLARATION.DecVictimeAgenceCodePostal,
  TBLDECLARATION.DecVictimeAgenceVille,
  TBLDECLARATION.DecVictimeAgencePays,
  TBLMARCHE.MarNom AS DecVictimeMarNom,
  TBLTYPEBENEFICIAIRE.TypBenNom AS DecBeneficiaireTypBenNom,
  TBLDECLARATION.DecBeneficiaireNom,
  TBLDECLARATION.DecBeneficiairePrenom,
  TBLDECLARATION.DecBeneficiaireNaissanceDate,
  TBLDECLARATION.DecBeneficiaireNaissanceLieu,
  TBLDECLARATION.DecBeneficiaireAdresse,
  TBLDECLARATION.DecBeneficiaireCodePostal,
  TBLDECLARATION.DecBeneficiaireVille,
  TBLDECLARATION.DecBeneficiairePays,
  TBLDECLARATION.DecBeneficiaireTelephone1,
  TBLDECLARATION.DecBeneficiaireTelephone2,
  TBLDECLARATION.DecBeneficiaireRIB,
  TBLDECLARATION.DecBeneficiaireBanque,
  TBLDECLARATION.DecBeneficiaireAgence,
  TBLDECLARATION.DecBeneficiaireAgenceAdresse,
  TBLDECLARATION.DecBeneficiaireAgenceCodePostal,
  TBLDECLARATION.DecBeneficiaireAgenceVille,
  TBLDECLARATION.DecBeneficiaireAgencePays,
  TBLDECLARATION.DecBeneficiaireAgenceAppel,
  TBLDECLARATION.DecBeneficiaireAgenceResponsableNom,
  TBLDECLARATION.DecBeneficiaireAgenceResponsablePrenom,
  TBLDECLARATION.DecBeneficiaireAgenceResponsableTelephone1,
  TBLDECLARATION.DecBeneficiaireAgenceResponsableTelephone2,
  TBLDECLARATION.DecBeneficiaireCommentaire,
  TBLDECLARATION.DecFraudeDate,
  TBLDECLARATION.DecFraudeHeure,
  TBLTYPEOPERATION.TypOpeNom AS DecFraudeTypOpeNom,
  TBLDECLARATION.DecFraudeMontant,
  TBLDECLARATION.DecFraudeMontantRecuperer,
  TBLDECLARATION.DecFraudeLibelle,
  TBLTYPEFRAUDE.TypFrdNom AS DecFraudeTypFrdNom,
  TBLSOURCEIP.SouIPNom AS DecFraudeSouIPNom,
  TBLDECLARATION.DecFraudeAdresseIP,
  TBLDECLARATION.DecFraudeDecaissementDate,
  TBLDECLARATION.DecFraudeDecaissementHeure,
  TBLDECLARATION.DecFraudeDecaissementMontant,
  TBLDECLARATION.DecFraudeCommentaire,
  TBLDECLARATION.DecDepotPlainteClient,
  TBLDECLARATION.DecDepotPlainteClientDate,
  TBLAUTORITE.AutNom AS DecDepotPlainteClientAutNom,
  TBLDECLARATION.DecDepotPlainteClientLieu,
  TBLDECLARATION.DecDepotPlainteClientAccepter,
  TBLDECLARATION.DecDepotPlainteClientMotif,
  TBLDECLARATION.DecDepotPlainteEntite,
  TBLDECLARATION.DecDepotPlainteEntiteDate,
  TBLAUTORITE_1.AutNom AS DecDepotPlainteEntiteAutNom,
  TBLDECLARATION.DecDepotPlainteEntiteLieu,
  TBLDECLARATION.DecDepotPlainteEntiteAccepter,
  TBLDECLARATION.DecDepotPlainteEntiteMotif,
  TBLDECLARATION.DecSiteIlliciteDate,
  TBLSTATUTSITE.StaSitNom AS DecSiteIlliciteStaSitNom,
  TBLSOURCESITE.SouSitNom AS DecSiteIlliciteSouSitNom,
  TBLDECLARATION.DecSiteIlliciteHebergeur,
  TBLDECLARATION.DecSiteIlliciteHebergeurAdresse,
  TBLDECLARATION.DecSiteIlliciteHebergeurCodePostal,
  TBLDECLARATION.DecSiteIlliciteHebergeurVille,
  TBLDECLARATION.DecSiteIlliciteHebergeurPays,
  TBLDECLARATION.DecSiteIlliciteSuivi,
  TBLDECLARATION.DecSiteIlliciteCommentaire,
  TBLDECLARATION.DecLiaisonSFDate,
  TBLDECLARATION.DecLiaisonSFRetour
FROM
  TBLTYPEOPERATION
  RIGHT JOIN (
    TBLTYPEDECLARATION
    RIGHT JOIN (
      TBLTYPECOMPTE
      RIGHT JOIN (
        TBLTYPEBENEFICIAIRE
        RIGHT JOIN (
          TBLSTATUTSITE
          RIGHT JOIN (
            TBLSOURCESITE
            RIGHT JOIN (
              TBLTYPEFRAUDE
              RIGHT JOIN (
                TBLMARCHE
                RIGHT JOIN (
                  TBLLOCALISATIONSITE
                  RIGHT JOIN (
                    TBLSOURCEIP
                    RIGHT JOIN (
                      TBLENTITE
                      RIGHT JOIN (
                        TBLAUTORITE
                        RIGHT JOIN (
                          TBLDECLARATION
                          LEFT JOIN TBLAUTORITE AS TBLAUTORITE_1 ON TBLDECLARATION.DecDepotPlainteEntiteAutNumero = TBLAUTORITE_1.AutNumero
                        ) ON TBLAUTORITE.AutNumero = TBLDECLARATION.DecDepotPlainteClientAutNumero
                      ) ON TBLENTITE.EntNumero = TBLDECLARATION.DecDeclarantEntNumero
                    ) ON TBLSOURCEIP.SouIPNumero = TBLDECLARATION.DecFraudeSouIPNumero
                  ) ON TBLLOCALISATIONSITE.LocSitNumero = TBLDECLARATION.DecSiteIlliciteLocSitNumero
                ) ON TBLMARCHE.MarNumero = TBLDECLARATION.DecVictimeMarNumero
              ) ON TBLTYPEFRAUDE.TypFrdNumero = TBLDECLARATION.DecFraudeTypFrdNumero
            ) ON TBLSOURCESITE.SouSitNumero = TBLDECLARATION.DecSiteIlliciteSouSitNumero
          ) ON TBLSTATUTSITE.StaSitNumero = TBLDECLARATION.DecSiteIlliciteStaSitNumero
        ) ON TBLTYPEBENEFICIAIRE.TypBenNumero = TBLDECLARATION.DecBeneficiaireTypBenNumero
      ) ON TBLTYPECOMPTE.TypCptNumero = TBLDECLARATION.DecVictimeTypCptNumero
    ) ON TBLTYPEDECLARATION.TypDecNumero = TBLDECLARATION.DecTypDecNumero
  ) ON TBLTYPEOPERATION.TypOpeNumero = TBLDECLARATION.DecFraudeTypOpeNumero
ORDER BY
  TBLDECLARATION.DecNumero;
