USE [master]
GO
/****** Object:  Database [Local_DB]    Script Date: 08/29/17 3:25:32 PM ******/
IF EXISTS(select * from sys.databases where name='Local_DB')
DROP DATABASE Local_DB
GO
CREATE DATABASE [Local_DB]
GO
ALTER DATABASE [Local_DB] SET COMPATIBILITY_LEVEL = 110
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [Local_DB].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [Local_DB] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [Local_DB] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [Local_DB] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [Local_DB] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [Local_DB] SET ARITHABORT OFF 
GO
ALTER DATABASE [Local_DB] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [Local_DB] SET AUTO_CREATE_STATISTICS ON 
GO
ALTER DATABASE [Local_DB] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [Local_DB] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [Local_DB] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [Local_DB] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [Local_DB] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [Local_DB] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [Local_DB] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [Local_DB] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [Local_DB] SET  DISABLE_BROKER 
GO
ALTER DATABASE [Local_DB] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [Local_DB] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [Local_DB] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [Local_DB] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [Local_DB] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [Local_DB] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [Local_DB] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [Local_DB] SET RECOVERY FULL 
GO
ALTER DATABASE [Local_DB] SET  MULTI_USER 
GO
ALTER DATABASE [Local_DB] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [Local_DB] SET DB_CHAINING OFF 
GO
ALTER DATABASE [Local_DB] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [Local_DB] SET TARGET_RECOVERY_TIME = 0 SECONDS 
GO
EXEC sys.sp_db_vardecimal_storage_format N'Millennium_SPEC_DB', N'ON'
GO
USE [Local_DB]
GO

GO
/****** Object:  DatabaseRole [readwrite]    Script Date: 08/29/17 3:25:33 PM ******/
CREATE ROLE [readwrite]
GO
/****** Object:  Table [dbo].[ANALYST]    Script Date: 08/29/17 3:25:34 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ANALYST](
	[Analyst_ID] [int] IDENTITY(1,1) NOT NULL,
	[First_Name] [nvarchar](50) NOT NULL,
	[Last_Name] [nvarchar](50) NULL,
	[Is_Analyst] [bit] NOT NULL,
	[Username] [nvarchar](50) NULL,
PRIMARY KEY CLUSTERED 
(
	[Analyst_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[DISCIPLINE]    Script Date: 08/29/17 3:25:34 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DISCIPLINE](
	[Discipline_ID] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[Discipline] [nvarchar](50) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[Discipline_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[SPEC]    Script Date: 08/29/17 3:25:34 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SPEC](
	[SPEC_ID] [int] IDENTITY(1,1) NOT NULL,
	[DEPARTMENT] [nvarchar](255) NULL,
	[SUMMARY] [nvarchar](255) NOT NULL,
	[DESCRIPTION] [nvarchar](max) NULL,
	[ANALYST] [nvarchar](50) NULL,
	[DATE_SUBMITTED] [date] NOT NULL,
	[DATE_COMPLETED] [date] NULL,
	[STATUS] [nvarchar](255) NOT NULL,
	[VALUE_TO_BUSINESS] [nvarchar](255) NULL,
	[DATE_STARTED] [date] NULL,
	[RANK] [int] NULL,
	[DISCIPLINE] [nvarchar](255) NULL,
	[CONTACT_NAME] [nvarchar](255) NULL,
	[CONTACT_INFO] [nvarchar](255) NULL,
PRIMARY KEY CLUSTERED 
(
	[SPEC_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[UPDATE]    Script Date: 08/29/17 3:25:34 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[UPDATE](
	[UPDATE_ID] [int] IDENTITY(2,1) NOT NULL,
	[UPDATE_DESC] [nvarchar](max) NOT NULL,
	[UPDATE_DATE] [date] NOT NULL,
	[UPDATE_ANALYST] [nvarchar](255) NULL,
	[SPEC_ID] [int] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[UPDATE_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  View [dbo].[SHAREPOINT]    Script Date: 08/29/17 3:25:34 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[SHAREPOINT]
AS
SELECT        t1.SPEC_ID, t1.RANK, t1.STATUS, t1.DISCIPLINE, t1.DEPARTMENT, t1.SUMMARY, t1.DESCRIPTION, t2.UPDATE_DATE, t2.UPDATE_DESC AS LATEST_UPDATE, 
                         t1.ANALYST, t1.DATE_SUBMITTED, t1.DATE_STARTED, t1.DATE_COMPLETED, t1.VALUE_TO_BUSINESS, t1.CONTACT_NAME, T1.CONTACT_INFO
FROM            SPEC t1 CROSS APPLY
                             (SELECT        TOP 1 UPDATE_DESC, UPDATE_DATE
                               FROM            [UPDATE] t2
                               WHERE        t2.SPEC_ID = t1.SPEC_ID
                               ORDER BY UPDATE_DATE DESC) t2
UNION
SELECT        SPEC.SPEC_ID, SPEC.RANK, SPEC.STATUS, SPEC.DISCIPLINE, SPEC.DEPARTMENT, SPEC.SUMMARY, SPEC.DESCRIPTION,
                             (SELECT        CASE WHEN UPDATE_DATE = '1900-01-01 00:00:00.000' THEN '' ELSE CONVERT(VARCHAR(10), UPDATE_DATE, 103) END),
                             (SELECT        'No Updates'), SPEC.ANALYST, SPEC.DATE_SUBMITTED, SPEC.DATE_STARTED, SPEC.DATE_COMPLETED, SPEC.VALUE_TO_BUSINESS, 
                         SPEC.CONTACT_NAME, SPEC.CONTACT_INFO
FROM            SPEC LEFT JOIN
                         [UPDATE] ON SPEC.SPEC_ID = [UPDATE].SPEC_ID
WHERE        [UPDATE].SPEC_ID IS NULL

GO
ALTER TABLE [dbo].[UPDATE]  WITH CHECK ADD  CONSTRAINT [FK_UPDATE_SPEC] FOREIGN KEY([SPEC_ID])
REFERENCES [dbo].[SPEC] ([SPEC_ID])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[UPDATE] CHECK CONSTRAINT [FK_UPDATE_SPEC]
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SHAREPOINT'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SHAREPOINT'
GO
USE [master]
GO
ALTER DATABASE [Local_DB] SET  READ_WRITE 
GO
