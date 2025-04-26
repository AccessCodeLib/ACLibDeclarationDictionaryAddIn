CREATE TABLE [tabWords] (
  [Word] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Variations] VARCHAR (255),
  [Diff] BIT
)
