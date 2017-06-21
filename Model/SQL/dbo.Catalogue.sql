USE [MetalSpec]
GO

/****** Object: Table [dbo].[Catalogue] Script Date: 09.02.2016 16:55:22 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Catalogue] (
    [ID]          INT          IDENTITY (1, 1) NOT NULL,
    [Name]        NCHAR (200)  NULL,
    [Description] NCHAR (1000) NULL
);


