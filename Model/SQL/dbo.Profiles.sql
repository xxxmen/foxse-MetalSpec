CREATE TABLE [dbo].[Profiles] (
    [ID]          INT            IDENTITY (1, 1) NOT NULL,
    [Name]        NVARCHAR (100) NULL,
    [h]           REAL           NULL,
    [b]           REAL           NULL,
    [s]           REAL           NULL,
    [t]           REAL           NULL,
    [r1]          REAL           NULL,
    [r2]          REAL           NULL,
    [A]           REAL           NULL,
    [ly]          REAL           NULL,
    [lz]          REAL           NULL,
    [lu]          REAL           NULL,
    [Wy]          REAL           NULL,
    [iy]          REAL           NULL,
    [lv]          REAL           NULL,
    [Wz]          REAL           NULL,
    [Wv]          REAL           NULL,
    [zo]          REAL           NULL,
    [tgAlpha]     REAL           NULL,
    [Sy]          REAL           NULL,
    [gamma]       REAL           NULL,
    [D]           REAL           NULL,
    [n1]          REAL           NULL,
    [n2]          REAL           NULL,
    [z0]          REAL           NULL,
    [Sz]          REAL           NULL,
    [h_2t]        REAL           NULL,
    [Wply]        REAL           NULL,
    [Wplz]        REAL           NULL,
    [Wvo]         REAL           NULL,
    [iz]          REAL           NULL,
    [iu]          REAL           NULL,
    [iv]          REAL           NULL,
    [lyz]         REAL           NULL,
    [yo]          REAL           NULL,
    [P]           REAL           NULL,
    [SortamentID] INT            NOT NULL,
    [PaintArea] REAL NULL, 
    CONSTRAINT [PK_Profiles] PRIMARY KEY CLUSTERED ([ID] ASC)
);


GO
CREATE NONCLUSTERED INDEX [IX_FK_SortamentProfiles]
    ON [dbo].[Profiles]([SortamentID] ASC);

