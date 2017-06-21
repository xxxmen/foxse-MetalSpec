
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 06/21/2017 23:35:45
-- Generated from EDMX file: C:\Users\fox\Documents\Visual Studio 2015\Projects\MetalSpec\Model\ProfilesModel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [MetalSpec];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[FK_CatalogueSortament]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Sortament] DROP CONSTRAINT [FK_CatalogueSortament];
GO
IF OBJECT_ID(N'[dbo].[FK_SortamentProfiles]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Profiles] DROP CONSTRAINT [FK_SortamentProfiles];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[Catalogue]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Catalogue];
GO
IF OBJECT_ID(N'[dbo].[Profiles]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Profiles];
GO
IF OBJECT_ID(N'[dbo].[Sortament]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Sortament];
GO
IF OBJECT_ID(N'[MetalSpecModelStoreContainer].[FireResistTypes]', 'U') IS NOT NULL
    DROP TABLE [MetalSpecModelStoreContainer].[FireResistTypes];
GO
IF OBJECT_ID(N'[MetalSpecModelStoreContainer].[Materials]', 'U') IS NOT NULL
    DROP TABLE [MetalSpecModelStoreContainer].[Materials];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'Catalogue'
CREATE TABLE [dbo].[Catalogue] (
    [ID] int IDENTITY(1,1) NOT NULL,
    [Name] nchar(200)  NULL,
    [Description] nchar(1000)  NULL
);
GO

-- Creating table 'Sortament'
CREATE TABLE [dbo].[Sortament] (
    [ID] int IDENTITY(1,1) NOT NULL,
    [Name] nchar(200)  NULL,
    [Description] nchar(1000)  NULL,
    [CatalogueID] int  NOT NULL
);
GO

-- Creating table 'Profiles'
CREATE TABLE [dbo].[Profiles] (
    [ID] int IDENTITY(1,1) NOT NULL,
    [Name] nchar(50)  NULL,
    [h] real  NULL,
    [b] real  NULL,
    [s] real  NULL,
    [t] real  NULL,
    [r1] real  NULL,
    [r2] real  NULL,
    [A] real  NULL,
    [ly] real  NULL,
    [lz] real  NULL,
    [lu] real  NULL,
    [Wy] real  NULL,
    [iy] real  NULL,
    [lv] real  NULL,
    [Wz] real  NULL,
    [Wv] real  NULL,
    [zo] real  NULL,
    [tgAlpha] real  NULL,
    [Sy] real  NULL,
    [gamma] real  NULL,
    [D] real  NULL,
    [n1] real  NULL,
    [n2] real  NULL,
    [z0] real  NULL,
    [Sz] real  NULL,
    [h_2t] real  NULL,
    [Wply] real  NULL,
    [Wplz] real  NULL,
    [Wvo] real  NULL,
    [iz] real  NULL,
    [iu] real  NULL,
    [iv] real  NULL,
    [lyz] real  NULL,
    [yo] real  NULL,
    [P] real  NULL,
    [SortamentID] int  NOT NULL,
    [PaintArea] real  NULL
);
GO

-- Creating table 'Materials'
CREATE TABLE [dbo].[Materials] (
    [Id] int  NOT NULL,
    [Name] nchar(50)  NULL,
    [Density] real  NULL,
    [Description] nchar(1000)  NULL
);
GO

-- Creating table 'FireResistTypes'
CREATE TABLE [dbo].[FireResistTypes] (
    [ID] int IDENTITY(1,1) NOT NULL,
    [Name] nchar(20)  NULL,
    [Description] nchar(160)  NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [ID] in table 'Catalogue'
ALTER TABLE [dbo].[Catalogue]
ADD CONSTRAINT [PK_Catalogue]
    PRIMARY KEY CLUSTERED ([ID] ASC);
GO

-- Creating primary key on [ID] in table 'Sortament'
ALTER TABLE [dbo].[Sortament]
ADD CONSTRAINT [PK_Sortament]
    PRIMARY KEY CLUSTERED ([ID] ASC);
GO

-- Creating primary key on [ID] in table 'Profiles'
ALTER TABLE [dbo].[Profiles]
ADD CONSTRAINT [PK_Profiles]
    PRIMARY KEY CLUSTERED ([ID] ASC);
GO

-- Creating primary key on [Id] in table 'Materials'
ALTER TABLE [dbo].[Materials]
ADD CONSTRAINT [PK_Materials]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [ID] in table 'FireResistTypes'
ALTER TABLE [dbo].[FireResistTypes]
ADD CONSTRAINT [PK_FireResistTypes]
    PRIMARY KEY CLUSTERED ([ID] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [SortamentID] in table 'Profiles'
ALTER TABLE [dbo].[Profiles]
ADD CONSTRAINT [FK_SortamentProfiles]
    FOREIGN KEY ([SortamentID])
    REFERENCES [dbo].[Sortament]
        ([ID])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_SortamentProfiles'
CREATE INDEX [IX_FK_SortamentProfiles]
ON [dbo].[Profiles]
    ([SortamentID]);
GO

-- Creating foreign key on [CatalogueID] in table 'Sortament'
ALTER TABLE [dbo].[Sortament]
ADD CONSTRAINT [FK_CatalogueSortament]
    FOREIGN KEY ([CatalogueID])
    REFERENCES [dbo].[Catalogue]
        ([ID])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_CatalogueSortament'
CREATE INDEX [IX_FK_CatalogueSortament]
ON [dbo].[Sortament]
    ([CatalogueID]);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------