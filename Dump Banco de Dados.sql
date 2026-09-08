USE [master]
GO
/****** Object:  Database [XYZCartoes]    Script Date: 08/09/2026 08:50:20 ******/
CREATE DATABASE [XYZCartoes]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'XYZCartoes', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL17.MSSQLSERVER\MSSQL\DATA\XYZCartoes.mdf' , SIZE = 8192KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'XYZCartoes_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL17.MSSQLSERVER\MSSQL\DATA\XYZCartoes_log.ldf' , SIZE = 8192KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
 WITH CATALOG_COLLATION = DATABASE_DEFAULT, LEDGER = OFF
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [XYZCartoes].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [XYZCartoes] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [XYZCartoes] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [XYZCartoes] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [XYZCartoes] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [XYZCartoes] SET ARITHABORT OFF 
GO
ALTER DATABASE [XYZCartoes] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [XYZCartoes] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [XYZCartoes] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [XYZCartoes] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [XYZCartoes] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [XYZCartoes] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [XYZCartoes] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [XYZCartoes] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [XYZCartoes] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [XYZCartoes] SET  ENABLE_BROKER 
GO
ALTER DATABASE [XYZCartoes] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [XYZCartoes] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [XYZCartoes] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [XYZCartoes] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [XYZCartoes] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [XYZCartoes] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [XYZCartoes] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [XYZCartoes] SET RECOVERY FULL 
GO
ALTER DATABASE [XYZCartoes] SET  MULTI_USER 
GO
ALTER DATABASE [XYZCartoes] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [XYZCartoes] SET DB_CHAINING OFF 
GO
ALTER DATABASE [XYZCartoes] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [XYZCartoes] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO
ALTER DATABASE [XYZCartoes] SET DELAYED_DURABILITY = DISABLED 
GO
ALTER DATABASE [XYZCartoes] SET ACCELERATED_DATABASE_RECOVERY = OFF  
GO
ALTER DATABASE [XYZCartoes] SET QUERY_STORE = ON
GO
ALTER DATABASE [XYZCartoes] SET QUERY_STORE (OPERATION_MODE = READ_WRITE, CLEANUP_POLICY = (STALE_QUERY_THRESHOLD_DAYS = 30), DATA_FLUSH_INTERVAL_SECONDS = 900, INTERVAL_LENGTH_MINUTES = 60, MAX_STORAGE_SIZE_MB = 1000, QUERY_CAPTURE_MODE = AUTO, SIZE_BASED_CLEANUP_MODE = AUTO, MAX_PLANS_PER_QUERY = 200, WAIT_STATS_CAPTURE_MODE = ON)
GO
USE [XYZCartoes]
GO
/****** Object:  UserDefinedFunction [dbo].[fn_CategoriaTransacao]    Script Date: 08/09/2026 08:50:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE   FUNCTION [dbo].[fn_CategoriaTransacao]
(
    @Valor DECIMAL(18, 2)
)
RETURNS VARCHAR(20)
AS
BEGIN
    DECLARE @Categoria VARCHAR(20);

    SET @Categoria =
        CASE
            WHEN @Valor > 2000 THEN 'Premium'
            WHEN @Valor >= 1000 AND @Valor <= 2000 THEN 'Alta'
            WHEN @Valor >= 500 AND @Valor < 1000 THEN 'Média'
            ELSE 'Baixa'
        END;

    RETURN @Categoria;
END;
GO
/****** Object:  Table [dbo].[Transacoes]    Script Date: 08/09/2026 08:50:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Transacoes](
	[Id_Transacao] [int] IDENTITY(1,1) NOT NULL,
	[Numero_Cartao] [varchar](16) NOT NULL,
	[Valor_Transacao] [decimal](18, 2) NOT NULL,
	[Data_Transacao] [datetime] NOT NULL,
	[Descricao] [varchar](255) NULL,
	[Status_Transacao] [varchar](20) NOT NULL,
 CONSTRAINT [UQ_Transacoes_Id_Transacao] UNIQUE NONCLUSTERED 
(
	[Id_Transacao] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  UserDefinedFunction [dbo].[fn_TransacoesCategorizadasPeriodo]    Script Date: 08/09/2026 08:50:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE   FUNCTION [dbo].[fn_TransacoesCategorizadasPeriodo]
(
    @Data_Inicial DATETIME,
    @Data_Final DATETIME
)
RETURNS TABLE
AS
RETURN
(
    SELECT
        Id_Transacao,
        Numero_Cartao,
        Data_Transacao,
        Valor_Transacao,
        Status_Transacao,
        Descricao,
        dbo.fn_CategoriaTransacao(Valor_Transacao) AS Categoria
    FROM dbo.Transacoes
    WHERE Data_Transacao >= @Data_Inicial
      AND Data_Transacao < DATEADD(DAY, 1, @Data_Final)
);
GO
/****** Object:  View [dbo].[vw_TransacoesFinanceiras]    Script Date: 08/09/2026 08:50:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE   VIEW [dbo].[vw_TransacoesFinanceiras]
AS
SELECT
    Id_Transacao,
    Numero_Cartao,
    Data_Transacao,
    YEAR(Data_Transacao) AS Ano,
    MONTH(Data_Transacao) AS Mes,
    Valor_Transacao,
    Status_Transacao,
    Descricao,
    dbo.fn_CategoriaTransacao(Valor_Transacao) AS Categoria
FROM dbo.Transacoes;
GO
/****** Object:  Table [dbo].[cboTransacoes]    Script Date: 08/09/2026 08:50:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[cboTransacoes](
	[Id] [int] NULL,
	[cboNome] [varchar](30) NOT NULL,
	[Descricao] [varchar](100) NOT NULL,
 CONSTRAINT [UQ_cboTransacoes_cboNome_Id] UNIQUE NONCLUSTERED 
(
	[cboNome] ASC,
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Transacoes_Erro]    Script Date: 08/09/2026 08:50:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Transacoes_Erro](
	[Id_Erro] [int] IDENTITY(1,1) NOT NULL,
	[Id_Transacao] [int] NULL,
	[Operacao] [varchar](20) NOT NULL,
	[Numero_Erro] [int] NULL,
	[Descricao_Erro] [varchar](4000) NULL,
	[Usuario_Log] [sysname] NOT NULL,
	[Data_Erro] [datetime] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[Id_Erro] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Transacoes_Log]    Script Date: 08/09/2026 08:50:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Transacoes_Log](
	[Id_Log] [int] IDENTITY(1,1) NOT NULL,
	[Id_Transacao] [int] NULL,
	[Operacao] [varchar](10) NOT NULL,
	[Numero_Cartao] [varchar](16) NULL,
	[Valor_Transacao] [money] NULL,
	[Data_Transacao] [datetime] NULL,
	[Descricao] [varchar](255) NULL,
	[Status_Transacao] [varchar](20) NULL,
	[Data_Log] [datetime] NOT NULL,
	[Usuario_Log] [sysname] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[Id_Log] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Transacoes_Teste_Bloqueio]    Script Date: 08/09/2026 08:50:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Transacoes_Teste_Bloqueio](
	[Id_Bloqueio] [int] IDENTITY(1,1) NOT NULL,
	[Id_Transacao] [int] NOT NULL,
	[Observacao] [varchar](100) NULL,
PRIMARY KEY CLUSTERED 
(
	[Id_Bloqueio] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_Transacoes_Periodo_Status_Cartao]    Script Date: 08/09/2026 08:50:20 ******/
CREATE NONCLUSTERED INDEX [IX_Transacoes_Periodo_Status_Cartao] ON [dbo].[Transacoes]
(
	[Data_Transacao] ASC,
	[Status_Transacao] ASC,
	[Numero_Cartao] ASC
)
INCLUDE([Valor_Transacao]) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Transacoes_Erro] ADD  DEFAULT (suser_sname()) FOR [Usuario_Log]
GO
ALTER TABLE [dbo].[Transacoes_Erro] ADD  DEFAULT (getdate()) FOR [Data_Erro]
GO
ALTER TABLE [dbo].[Transacoes_Log] ADD  DEFAULT (getdate()) FOR [Data_Log]
GO
ALTER TABLE [dbo].[Transacoes_Log] ADD  DEFAULT (suser_sname()) FOR [Usuario_Log]
GO
ALTER TABLE [dbo].[Transacoes_Teste_Bloqueio]  WITH CHECK ADD  CONSTRAINT [FK_Transacoes_Teste_Bloqueio_Transacoes] FOREIGN KEY([Id_Transacao])
REFERENCES [dbo].[Transacoes] ([Id_Transacao])
GO
ALTER TABLE [dbo].[Transacoes_Teste_Bloqueio] CHECK CONSTRAINT [FK_Transacoes_Teste_Bloqueio_Transacoes]
GO
/* Dados iniciais para testes */
INSERT INTO dbo.Transacoes
    (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao)
VALUES
    ('1234567890123456', 150.90, '2026-08-01T09:15:00', 'Compra teste', 'Aprovada'),
    ('4111111111111111', 125.90, '2026-08-03T10:20:00', 'Compra supermercado', 'Aprovada'),
    ('5500000000000004', 89.50,  '2026-08-05T14:35:00', 'Posto de combustivel', 'Pendente'),
    ('2223000048400037', 250.00, '2026-08-08T16:40:00', 'Parcela eletrodomestico', 'Aprovada'),
    ('6011000000000004', 42.75,  '2026-08-10T08:05:00', 'Farmacia', 'Cancelada'),
    ('340000000000009', 315.20,  '2026-08-12T12:10:00', 'Restaurante', 'Aprovada'),
    ('30000000000004',  67.30,   '2026-08-15T18:25:00', 'Aplicativo de transporte', 'Pendente'),
    ('2223000048400086', 980.00, '2026-08-18T11:50:00', 'Hospedagem hotel', 'Aprovada'),
    ('4000000000000002', 154.10, '2026-08-20T13:45:00', 'Loja de roupas', 'Cancelada'),
    ('4000000000000005', 2390.00, '2026-08-22T15:30:00', 'Material de construcao', 'Aprovada'),
    ('5555555555554444', 75.40,  '2026-08-25T09:05:00', 'Livraria', 'Aprovada'),
    ('378282246310005', 1100.00, '2026-08-27T17:15:00', 'Curso profissionalizante', 'Pendente'),
    ('6011111111111117', 520.00, '2026-08-29T19:00:00', 'Manutencao veicular', 'Aprovada'),
    ('3530111333300000', 1999.99, '2026-09-01T08:30:00', 'Pagamento de servico', 'Aprovada'),
    ('2221000000000009', 480.00, '2026-09-02T10:45:00', 'Consulta medica', 'Cancelada'),
    ('3566002020360505', 2050.00, '2026-09-03T14:20:00', 'Compra de equipamento', 'Pendente'),
    ('30569309025904',  35.90,   '2026-09-04T16:05:00', 'Lanchonete', 'Aprovada'),
    ('38520000023237',   750.00, '2026-09-05T18:10:00', 'Seguro residencial', 'Aprovada'),
    ('6011000990139424', 3200.00, '2026-09-06T11:25:00', 'Pagamento premium', 'Aprovada'),
    ('4222222222222222', 645.80, '2026-09-07T20:40:00', 'Compra online', 'Pendente');
GO
/****** Object:  StoredProcedure [dbo].[sp_TotalTransacoesPeriodo]    Script Date: 08/09/2026 08:50:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE   PROCEDURE [dbo].[sp_TotalTransacoesPeriodo]
    @Data_Inicial DATETIME,
    @Data_Final DATETIME,
    @Status_Transacao VARCHAR(20) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        Numero_Cartao,
        SUM(Valor_Transacao) AS Valor_Total,
        COUNT(*) AS Quantidade_Transacoes,
        Status_Transacao
    FROM dbo.Transacoes
    WHERE Data_Transacao >= @Data_Inicial
      AND Data_Transacao < DATEADD(DAY, 1, @Data_Final)
      AND (
            @Status_Transacao IS NULL
            OR @Status_Transacao = ''
            OR Status_Transacao = @Status_Transacao
          )
    GROUP BY
        Numero_Cartao,
        Status_Transacao
    ORDER BY
        Numero_Cartao,
        Status_Transacao;
END;

GO
USE [master]
GO
ALTER DATABASE [XYZCartoes] SET  READ_WRITE 
GO

