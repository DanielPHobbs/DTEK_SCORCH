USE [SCORCHPersistantDB]
GO

/****** Object:  Table [dbo].[StateTracking01]    Script Date: 17/10/2021 13:28:35 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[StateTracking01](
	[Guid] [varchar](255) NOT NULL,
	[RunCount] [int] NULL,
	[RunbookName] [varchar](255) NULL,
	[Description] [varchar](255) NULL,
	[HomedriveMap] [varchar](255) NULL,
	[ExchangeMB] [varchar](255) NULL,
	[GroupMembership] [varchar](255) NULL,
	[ActivityName] [varchar](255) NULL,
	[ActivityStatus] [datetime] NULL,
	[AvctivityStart] [datetime] NULL,
	[ActivityEnd] [datetime] NULL,
	[ErrorSeverity] [varchar](255) NULL,
	[ErrorReason] [varchar](255) NULL,
	[AllTaskStatus] [varchar](255) NULL,
	[RunbookServer] [varchar](255) NULL,
	[ExtendedData1] [varchar](255) NULL,
	[ExtendedData2] [varchar](255) NULL,
	[ExtendedData3] [varchar](255) NULL,
	[Displayname] [varchar](75) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[Guid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
