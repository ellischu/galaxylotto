USE [GL_DBLotto539]
GO

/****** Object:  Table [dbo].[tblMissAll01]    Script Date: 2016/11/12 下午 03:38:41 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tblMissAll01](
	[lngMiss01SN] [bigint] IDENTITY(1,1) NOT NULL,
	[lngTotalSN] [bigint] NOT NULL,
	[lngNums] [smallint] NULL,
	[strMissOption] [varchar](max) NULL,
	[lngExist] [bigint] NULL,
 CONSTRAINT [PK_tblMissAll01] PRIMARY KEY CLUSTERED 
(
	[lngMiss01SN] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [dbo].[tblMissAll01]  WITH CHECK ADD  CONSTRAINT [FK_tblMissAll01_tblData] FOREIGN KEY([lngTotalSN])
REFERENCES [dbo].[tblData] ([lngTotalSN])
GO

ALTER TABLE [dbo].[tblMissAll01] CHECK CONSTRAINT [FK_tblMissAll01_tblData]
GO

EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'gen及前8項欄位，遺漏值統計表' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblMissAll01'
GO