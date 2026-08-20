CREATE TABLE [TBLMENU] (
  [MenGroupe] LONG,
  [MenNumero] LONG,
  [MenNom] VARCHAR (100),
  [MenMacro] VARCHAR (250),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([MenGroupe], [MenNumero])
)
