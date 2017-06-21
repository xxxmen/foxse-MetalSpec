USE [MetalSpec]
GO

/****** Object: Table [dbo].[Materials] Script Date: 09.02.2016 13:31:23 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Materials] (
    [Id]          INT          IDENTITY (1, 1) NOT NULL,
    [Name]        NCHAR (50)  NULL,
    [Description] NCHAR (1000) NULL,
    [Density]     REAL        NULL
);