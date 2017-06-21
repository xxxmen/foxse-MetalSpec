USE [MetalSpec]
GO

/****** Object:  Table [dbo].[Profiles]    Script Date: 02/04/2016 15:11:49 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Profiles]') AND type in (N'U'))
DROP TABLE [dbo].[Profiles]
GO

USE [MetalSpec]
GO

/****** Object:  Table [dbo].[Profiles]    Script Date: 02/04/2016 15:11:49 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Profiles](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [nchar](50) NULL,
	[SortamentID] [int] NULL,
	[h] [real] NULL,
	[b] [real] NULL,
	[s] [real] NULL,
	[t] [real] NULL,
	[r1] [real] NULL,
	[r2] [real] NULL,
	[A] [real] NULL,
	[ly] [real] NULL,
	[lz] [real] NULL,
	[lu] [real] NULL,
	[Wy] [real] NULL,
	[iy] [real] NULL,
	[lv] [real] NULL,
	[Wz] [real] NULL,
	[Wv] [real] NULL,
	[zo] [real] NULL,
	[tgAlpha] [real] NULL,
	[Sy] [real] NULL,
	[gamma] [real] NULL,
	[D] [real] NULL,
	[n1] [real] NULL,
	[n2] [real] NULL,
	[z0] [real] NULL,
	[Sz] [real] NULL,
	[h-2t] [real] NULL,											
	[Wply] [real] NULL,
	[Wplz] [real] NULL,
	[Wvo] [real] NULL,
	[iz] [real] NULL,
	[iu] [real] NULL,
	[iv] [real] NULL,
	[lyz] [real] NULL,
	[yo] [real] NULL,
	[P] [real] NULL
) ON [PRIMARY]

GO

EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Профили' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Profiles'
GO

