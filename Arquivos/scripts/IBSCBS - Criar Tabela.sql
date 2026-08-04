/** Object:  Table [dbo].[TbIBSCBS]    Script Date: 23/04/2026 08:24:52 **/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('[dbo].[TbIBSCBS]') AND type = 'U')
BEGIN
    CREATE TABLE [dbo].[TbIBSCBS](
    	[CST] [varchar](3) NOT NULL,
    	[DescricaoIBSCBS] [varchar](255) NOT NULL,
    	[ind_gIBSCBS] [bit] NOT NULL,
    	[ind_gIBSCBSMono] [bit] NOT NULL,
    	[ind_gRed] [bit] NOT NULL,
    	[ind_gDif] [bit] NOT NULL,
    	[ind_gTransfCred] [bit] NOT NULL,
    	[ind_gCredPresIBSZFM] [bit] NOT NULL,
    	[ind_gAjusteCompet] [bit] NOT NULL,
    	[ind_RedutorBC] [bit] NOT NULL,
     CONSTRAINT [CSTIBSCBS] PRIMARY KEY CLUSTERED
    (
    	[CST] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
    ) ON [PRIMARY]
END
GO