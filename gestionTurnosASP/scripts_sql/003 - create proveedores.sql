CREATE TABLE [dbo].[proveedores](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[nombre] [varchar](200) NOT NULL,
	[telefono] [varchar](50) NULL,
	[celular] [varchar](50) NULL,
	[mail] [varchar](50) NULL,
	[created_by] [varchar](20) NOT NULL,
	[creation_date] [datetime] NOT NULL,
	[last_updated_by] [varchar](20) NOT NULL,
	[last_update_date] [datetime] NOT NULL,
	[empnro] [int] NOT NULL,
	[direccion] [varchar](100) NULL,
	[idciudad] [int] NULL
)