
CREATE TABLE [dbo].[cheques_estado](
	[id] [int] NOT NULL,
	[estado] [varchar](150) NOT NULL
) ON [PRIMARY]



INSERT [dbo].[cheques_estado] ([id], [estado]) VALUES (1, CONVERT(TEXT, N'Pagado'))
INSERT [dbo].[cheques_estado] ([id], [estado]) VALUES (2, CONVERT(TEXT, N'Entregado'))
INSERT [dbo].[cheques_estado] ([id], [estado]) VALUES (3, CONVERT(TEXT, N'Pendiente Entregar'))
INSERT [dbo].[cheques_estado] ([id], [estado]) VALUES (4, CONVERT(TEXT, N'Cobrado'))
INSERT [dbo].[cheques_estado] ([id], [estado]) VALUES (5, CONVERT(TEXT, N'Pendiente Entregar/Cobrar'))
INSERT [dbo].[cheques_estado] ([id], [estado]) VALUES (6, CONVERT(TEXT, N'Pendiente Asociar Venta'))
