USE [master]
GO
/****** Object:  Database [Nga_Siok_Thong_Sip_Ngoo_Im]    Script Date: 2024/10/1 下午 02:49:39 ******/
CREATE DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'Nga_Siok_Thong_Sip_Ngoo_Im', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\Nga_Siok_Thong_Sip_Ngoo_Im.mdf' , SIZE = 73728KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'Nga_Siok_Thong_Sip_Ngoo_Im_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\Nga_Siok_Thong_Sip_Ngoo_Im_log.ldf' , SIZE = 73728KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
 WITH CATALOG_COLLATION = DATABASE_DEFAULT, LEDGER = OFF
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET COMPATIBILITY_LEVEL = 160
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [Nga_Siok_Thong_Sip_Ngoo_Im].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET ARITHABORT OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET  DISABLE_BROKER 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET RECOVERY FULL 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET  MULTI_USER 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET DB_CHAINING OFF 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET DELAYED_DURABILITY = DISABLED 
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET ACCELERATED_DATABASE_RECOVERY = OFF  
GO
EXEC sys.sp_db_vardecimal_storage_format N'Nga_Siok_Thong_Sip_Ngoo_Im', N'ON'
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET QUERY_STORE = ON
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET QUERY_STORE (OPERATION_MODE = READ_WRITE, CLEANUP_POLICY = (STALE_QUERY_THRESHOLD_DAYS = 30), DATA_FLUSH_INTERVAL_SECONDS = 900, INTERVAL_LENGTH_MINUTES = 60, MAX_STORAGE_SIZE_MB = 1000, QUERY_CAPTURE_MODE = AUTO, SIZE_BASED_CLEANUP_MODE = AUTO, MAX_PLANS_PER_QUERY = 200, WAIT_STATS_CAPTURE_MODE = ON)
GO
USE [Nga_Siok_Thong_Sip_Ngoo_Im]
GO
/****** Object:  Table [dbo].[Han_Ji_Tian]    Script Date: 2024/10/1 下午 02:49:39 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Han_Ji_Tian](
	[識別號] [int] NOT NULL,
	[漢字] [nvarchar](50) NOT NULL,
	[聲母] [nvarchar](50) NOT NULL,
	[韻母] [nvarchar](50) NOT NULL,
	[聲調] [nvarchar](50) NOT NULL,
	[常用度] [float] NULL,
	[聲母識別碼] [int] NOT NULL,
	[韻母識別碼] [int] NOT NULL,
	[聲調識別碼] [int] NOT NULL,
	[column10] [nvarchar](1) NULL,
	[column11] [nvarchar](1) NULL,
	[column12] [nvarchar](1) NULL,
	[column13] [nvarchar](1) NULL,
	[column14] [nvarchar](1) NULL,
	[column15] [nvarchar](1) NULL,
 CONSTRAINT [PK_Han_Ji_Tian] PRIMARY KEY CLUSTERED 
(
	[識別號] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Siann_Bu_Piau]    Script Date: 2024/10/1 下午 02:49:39 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Siann_Bu_Piau](
	[識別號] [int] NOT NULL,
	[韻母編碼] [nvarchar](50) NOT NULL,
	[十五音字母] [nvarchar](50) NOT NULL,
	[韻母序] [tinyint] NOT NULL,
	[舒促] [nvarchar](50) NOT NULL,
	[國際音標] [nvarchar](50) NULL,
	[台語音標] [nvarchar](50) NULL,
	[方音符號] [nvarchar](50) NULL,
	[白話字] [nvarchar](50) NULL,
	[台羅拚音] [nvarchar](50) NULL,
	[閩拼] [nvarchar](50) NULL,
 CONSTRAINT [PK_Siann_Bu_Piau] PRIMARY KEY CLUSTERED 
(
	[識別號] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Siann_Tiau_Piau]    Script Date: 2024/10/1 下午 02:49:39 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Siann_Tiau_Piau](
	[識別號] [int] NOT NULL,
	[聲調] [nvarchar](50) NOT NULL,
	[四聲調] [nvarchar](50) NOT NULL,
	[舒促聲] [nvarchar](50) NOT NULL,
	[台羅八聲調] [tinyint] NOT NULL,
 CONSTRAINT [PK_Siann_Tiau_Piau] PRIMARY KEY CLUSTERED 
(
	[識別號] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Un_Bu_Piau]    Script Date: 2024/10/1 下午 02:49:39 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Un_Bu_Piau](
	[識別號] [int] NOT NULL,
	[韻母編碼] [nvarchar](50) NOT NULL,
	[十五音字母] [nvarchar](50) NOT NULL,
	[韻母序] [tinyint] NOT NULL,
	[舒促] [nvarchar](50) NOT NULL,
	[國際音標] [nvarchar](50) NULL,
	[台語音標] [nvarchar](50) NULL,
	[方音符號] [nvarchar](50) NULL,
	[白話字] [nvarchar](50) NULL,
	[台羅拚音] [nvarchar](50) NULL,
	[閩拼] [nvarchar](50) NULL,
 CONSTRAINT [PK_Un_Bu_Piau] PRIMARY KEY CLUSTERED 
(
	[識別號] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
USE [master]
GO
ALTER DATABASE [Nga_Siok_Thong_Sip_Ngoo_Im] SET  READ_WRITE 
GO
