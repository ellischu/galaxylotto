USE [GL_DBLotto539]
GO

/****** Object:  Table [dbo].[tblMissAll02]    Script Date: 2016/11/12 下午 04:32:53 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tblMissAll02](
	[lngMiss02SN] [bigint] IDENTITY(1,1) NOT NULL,
	[lngTotalSN] [bigint] NOT NULL,
	[lngNums] [smallint] NULL,
	[strMissOption] [varchar](max) NULL,
 CONSTRAINT [PK_tblMissAll02] PRIMARY KEY CLUSTERED 
(
	[lngMiss02SN] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [dbo].[tblMissAll02]  WITH CHECK ADD  CONSTRAINT [FK_tblMissAll02_tblData] FOREIGN KEY([lngTotalSN])
REFERENCES [dbo].[tblData] ([lngTotalSN])
GO

ALTER TABLE [dbo].[tblMissAll02] CHECK CONSTRAINT [FK_tblMissAll02_tblData]
GO

EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'gen及前8項欄位，出現數字遺漏值統計表' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblMissAll02'
GO


