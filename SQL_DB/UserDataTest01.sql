USE [SCORCHPersistantDB]
GO

/****** Object:  Table [dbo].[UserDataTest01]    Script Date: 28/07/2021 22:16:25 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[UserDataTest01](
	[First Name] [nvarchar](50) NULL,
	[Last Name] [nvarchar](50) NULL,
	[Gender] [nchar](10) NULL,
	[Country] [nchar](10) NULL,
	[Age] [numeric](18, 0) NULL,
	[Date] [datetime] NULL,
	[Id] [numeric](18, 0) NULL
) ON [PRIMARY]
GO
