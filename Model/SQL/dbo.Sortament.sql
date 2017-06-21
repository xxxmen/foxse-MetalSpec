USE [MetalSpec]
GO

/****** Object: Table [dbo].[Sortament] Script Date: 09.02.2016 16:55:40 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Sortament] (
    [ID]          INT          IDENTITY (1, 1) NOT NULL,
    [Name]        NCHAR (200)  NULL,
    [Description] NCHAR (1000) NULL,
    [CatalogueID] INT          NOT NULL
);


GO
CREATE NONCLUSTERED INDEX [IX_FK_CatalogueSortament]
    ON [dbo].[Sortament]([CatalogueID] ASC);


GO
ALTER TABLE [dbo].[Sortament]
    ADD CONSTRAINT [PK_Sortament] PRIMARY KEY CLUSTERED ([ID] ASC);


GO
ALTER TABLE [dbo].[Sortament]
    ADD CONSTRAINT [FK_CatalogueSortament] FOREIGN KEY ([CatalogueID]) REFERENCES [dbo].[Catalogue] ([ID]);


