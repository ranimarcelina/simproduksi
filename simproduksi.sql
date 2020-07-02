USE [master]
GO
/****** Object:  Database [simproduksi]    Script Date: 02/07/2020 19.42.34 ******/
CREATE DATABASE [simproduksi]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'simproduksi', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\DATA\simproduksi.mdf' , SIZE = 8192KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'simproduksi_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\DATA\simproduksi_log.ldf' , SIZE = 8192KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
GO
ALTER DATABASE [simproduksi] SET COMPATIBILITY_LEVEL = 140
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [simproduksi].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [simproduksi] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [simproduksi] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [simproduksi] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [simproduksi] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [simproduksi] SET ARITHABORT OFF 
GO
ALTER DATABASE [simproduksi] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [simproduksi] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [simproduksi] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [simproduksi] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [simproduksi] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [simproduksi] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [simproduksi] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [simproduksi] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [simproduksi] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [simproduksi] SET  DISABLE_BROKER 
GO
ALTER DATABASE [simproduksi] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [simproduksi] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [simproduksi] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [simproduksi] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [simproduksi] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [simproduksi] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [simproduksi] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [simproduksi] SET RECOVERY FULL 
GO
ALTER DATABASE [simproduksi] SET  MULTI_USER 
GO
ALTER DATABASE [simproduksi] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [simproduksi] SET DB_CHAINING OFF 
GO
ALTER DATABASE [simproduksi] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [simproduksi] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO
ALTER DATABASE [simproduksi] SET DELAYED_DURABILITY = DISABLED 
GO
EXEC sys.sp_db_vardecimal_storage_format N'simproduksi', N'ON'
GO
ALTER DATABASE [simproduksi] SET QUERY_STORE = OFF
GO
USE [simproduksi]
GO
/****** Object:  Table [dbo].[jadwalproduksi]    Script Date: 02/07/2020 19.42.34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[jadwalproduksi](
	[id_jadwalproduksi] [varchar](15) NOT NULL,
	[no_order] [varchar](15) NOT NULL,
	[customer] [varchar](100) NOT NULL,
	[product] [varchar](5) NOT NULL,
	[gsm] [int] NOT NULL,
	[size] [varchar](10) NOT NULL,
	[total] [int] NOT NULL,
	[cargo_ready] [date] NOT NULL,
	[date] [date] NOT NULL,
 CONSTRAINT [PK_jadwalproduksi] PRIMARY KEY CLUSTERED 
(
	[id_jadwalproduksi] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[mesinproduksi]    Script Date: 02/07/2020 19.42.34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[mesinproduksi](
	[kode_mesin] [varchar](6) NOT NULL,
	[nama_mesin] [varchar](20) NULL,
	[kapasitas] [varchar](50) NULL,
 CONSTRAINT [PK_mesinproduksi] PRIMARY KEY CLUSTERED 
(
	[kode_mesin] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pegawai]    Script Date: 02/07/2020 19.42.34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[pegawai](
	[id_pegawai] [nchar](6) NOT NULL,
	[nama_pegawai] [varchar](10) NOT NULL,
	[password] [int] NOT NULL,
	[status] [varchar](10) NOT NULL,
 CONSTRAINT [PK_pegawai] PRIMARY KEY CLUSTERED 
(
	[id_pegawai] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pemesanan]    Script Date: 02/07/2020 19.42.34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[pemesanan](
	[no_order] [varchar](15) NOT NULL,
	[customer] [varchar](100) NOT NULL,
	[product] [varchar](5) NOT NULL,
	[gsm] [int] NOT NULL,
	[size] [varchar](10) NOT NULL,
	[total] [int] NOT NULL,
	[cargo_ready] [date] NOT NULL,
	[date] [date] NOT NULL,
 CONSTRAINT [PK_pemesanan] PRIMARY KEY CLUSTERED 
(
	[no_order] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[produk]    Script Date: 02/07/2020 19.42.35 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[produk](
	[kode_produk] [varchar](6) NOT NULL,
	[nama_produk] [varchar](30) NOT NULL,
	[jenis_produk] [varchar](15) NOT NULL,
 CONSTRAINT [PK_produk] PRIMARY KEY CLUSTERED 
(
	[kode_produk] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
USE [master]
GO
ALTER DATABASE [simproduksi] SET  READ_WRITE 
GO
