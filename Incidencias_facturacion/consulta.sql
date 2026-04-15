-- 1. Crear el esquema si no existe
IF NOT EXISTS (SELECT * FROM sys.schemas WHERE name = 'Inventario')
BEGIN
    EXEC('CREATE SCHEMA Inventario')
END
GO

-- 2. Crear la tabla
CREATE TABLE [Inventario].[Codigos_QR] (
    [CodigoID]    NVARCHAR (50)  NOT NULL, -- Cambiado a NVARCHAR para aceptar "MTO0001"
    [Material]    NVARCHAR (MAX) NOT NULL,
    [Abreviatura] NVARCHAR (50)  NOT NULL,
    [Descripcion] NVARCHAR (MAX) NULL,
    [Empresa]     NVARCHAR (50)  NOT NULL,
    [Edificio]    NVARCHAR (50)  NOT NULL,
    [Area]        NVARCHAR (50)  NOT NULL,
    [RutaQR]      NVARCHAR (MAX) NULL,
    [RutaBarras]  NVARCHAR (MAX) NULL,
    [FechaHora]   DATETIME       NULL,
    CONSTRAINT [PK_Codigos_QR] PRIMARY KEY CLUSTERED ([CodigoID] ASC)
);
GO