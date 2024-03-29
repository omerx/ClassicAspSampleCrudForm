CREATE TABLE [dbo].[Table]
(
    [ID] [int] IDENTITY(1,1) NOT NULL,
    [VarcharCol] [varchar](50) NULL,
    [IntegerCol] [int] NULL,
    [DateCol] [datetime] NULL,
    CONSTRAINT [PK_TABLE] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO