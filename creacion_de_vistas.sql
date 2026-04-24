USE BI;
GO

IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'Inventario')
BEGIN
    EXEC('CREATE SCHEMA Inventario');
END
GO

PRINT 'Eliminando objetos anteriores...';
GO

IF OBJECT_ID('Inventario.vw_Otros_Equipos', 'V') IS NOT NULL DROP VIEW Inventario.vw_Otros_Equipos;
IF OBJECT_ID('Inventario.vw_CPU_Completo', 'V') IS NOT NULL DROP VIEW Inventario.vw_CPU_Completo;
GO

IF OBJECT_ID('Inventario.CPU_Relaciones', 'U') IS NOT NULL DROP TABLE Inventario.CPU_Relaciones;
IF OBJECT_ID('Inventario.CPU_Perifericos', 'U') IS NOT NULL DROP TABLE Inventario.CPU_Perifericos;
IF OBJECT_ID('Inventario.CPU_Software', 'U') IS NOT NULL DROP TABLE Inventario.CPU_Software;
IF OBJECT_ID('Inventario.Otros_Equipos', 'U') IS NOT NULL DROP TABLE Inventario.Otros_Equipos;
IF OBJECT_ID('Inventario.Auditoria', 'U') IS NOT NULL DROP TABLE Inventario.Auditoria;
IF OBJECT_ID('Inventario.CPU', 'U') IS NOT NULL DROP TABLE Inventario.CPU;
GO

PRINT 'Creando tablas nuevas...';
GO

CREATE TABLE Inventario.CPU (
    ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Host NVARCHAR(100) NOT NULL,
    No_Serie NVARCHAR(200) NULL,
    Empresa NVARCHAR(200) NULL,
    Edificio NVARCHAR(100) NULL,
    Area NVARCHAR(200) NULL,
    Estado NVARCHAR(50) NULL,
    Marca NVARCHAR(200) NULL,
    Modelo NVARCHAR(200) NULL,
    Procesador NVARCHAR(150) NULL,
    RAM NVARCHAR(50) NULL,
    Capacidad_Disco NVARCHAR(50) NULL,
    Tipo_Disco NVARCHAR(20) NULL,
    Observaciones NVARCHAR(MAX) NULL,
    Codigo_QR NVARCHAR(500) NULL,
    Codigo_Barras_CPU NVARCHAR(255) NULL,
    Timestamp DATETIME NULL,
    Fecha_Creacion DATETIME NOT NULL CONSTRAINT DF_CPU_FechaCreacion DEFAULT GETDATE(),
    Fecha_Modificacion DATETIME NOT NULL CONSTRAINT DF_CPU_FechaMod DEFAULT GETDATE(),
    CONSTRAINT UQ_CPU_Host UNIQUE (Host),
    CONSTRAINT UQ_CPU_CodigoBarras UNIQUE (Codigo_Barras_CPU),
    CONSTRAINT UQ_CPU_CodigoQR UNIQUE (Codigo_QR)
);
GO

CREATE TABLE Inventario.CPU_Software (
    ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Host_CPU NVARCHAR(100) NOT NULL,
    SO NVARCHAR(100) NULL,
    Office NVARCHAR(100) NULL,
    Antivirus NVARCHAR(100) NULL,
    Lector_PDF NVARCHAR(100) NULL,
    ERP NVARCHAR(100) NULL,
    Otro_1 NVARCHAR(100) NULL,
    Otro_2 NVARCHAR(100) NULL,
    Otro_3 NVARCHAR(100) NULL,
    Timestamp DATETIME NULL,
    Fecha_Creacion DATETIME NOT NULL CONSTRAINT DF_CPU_SW_FechaCreacion DEFAULT GETDATE(),
    Fecha_Modificacion DATETIME NOT NULL CONSTRAINT DF_CPU_SW_FechaMod DEFAULT GETDATE(),
    CONSTRAINT UQ_CPU_Software_Host UNIQUE (Host_CPU),
    CONSTRAINT FK_CPU_Software_Host FOREIGN KEY (Host_CPU)
        REFERENCES Inventario.CPU(Host)
        ON DELETE CASCADE
        ON UPDATE CASCADE
);
GO

CREATE TABLE Inventario.CPU_Perifericos (
    ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Host_CPU NVARCHAR(100) NOT NULL,
    Periferico_UID NVARCHAR(100) NOT NULL,
    Tipo NVARCHAR(50) NOT NULL,
    Modelo NVARCHAR(100) NULL,
    No_Serie NVARCHAR(200) NULL,
    Marca NVARCHAR(200) NULL,
    Estado NVARCHAR(50) NULL,
    Observaciones NVARCHAR(MAX) NULL,
    Codigo_Barras NVARCHAR(255) NULL,
    Codigo_ID NVARCHAR(100) NULL,
    Timestamp DATETIME NULL,
    Fecha_Creacion DATETIME NOT NULL CONSTRAINT DF_CPU_Per_FechaCreacion DEFAULT GETDATE(),
    Fecha_Modificacion DATETIME NOT NULL CONSTRAINT DF_CPU_Per_FechaMod DEFAULT GETDATE(),
    CONSTRAINT UQ_CPU_Perifericos_UID UNIQUE (Periferico_UID),
    CONSTRAINT UQ_CPU_Perifericos_CodigoID UNIQUE (Codigo_ID),
    CONSTRAINT FK_CPU_Perifericos_Host FOREIGN KEY (Host_CPU)
        REFERENCES Inventario.CPU(Host)
        ON DELETE CASCADE
        ON UPDATE CASCADE
);
GO

CREATE TABLE Inventario.CPU_Relaciones (
    ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Codigo_QR NVARCHAR(500) NULL,
    Codigo_Barras_CPU NVARCHAR(255) NULL,
    Tipo_Periferico NVARCHAR(50) NULL,
    Codigo_Barras_Periferico NVARCHAR(255) NULL,
    Timestamp DATETIME NULL,
    Fecha_Creacion DATETIME NOT NULL CONSTRAINT DF_CPU_Rel_FechaCreacion DEFAULT GETDATE()
);
GO

CREATE TABLE Inventario.Otros_Equipos (
    ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Tipo NVARCHAR(100) NOT NULL,
    Nombre NVARCHAR(500) NULL,
    No_Serie NVARCHAR(200) NULL,
    Marca NVARCHAR(200) NULL,
    Modelo NVARCHAR(200) NULL,
    Empresa NVARCHAR(200) NULL,
    Edificio NVARCHAR(100) NULL,
    Area NVARCHAR(200) NULL,
    Ubicacion_En_Edificio NVARCHAR(250) NULL,
    Estado NVARCHAR(50) NULL,
    Tipo_Sensor NVARCHAR(100) NULL,
    Resolucion_Pantalla NVARCHAR(50) NULL,
    Sistema_Operativo NVARCHAR(100) NULL,
    Observaciones NVARCHAR(MAX) NULL,
    Codigo_Barras NVARCHAR(255) NULL,
    Codigo_ID NVARCHAR(100) NULL,
    Timestamp DATETIME NULL,
    Fecha_Creacion DATETIME NOT NULL CONSTRAINT DF_Otros_FechaCreacion DEFAULT GETDATE(),
    Fecha_Modificacion DATETIME NOT NULL CONSTRAINT DF_Otros_FechaMod DEFAULT GETDATE(),
    CONSTRAINT UQ_Otros_CodigoID UNIQUE (Codigo_ID)
);
GO

CREATE TABLE Inventario.Auditoria (
    ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Tabla_Afectada NVARCHAR(100) NULL,
    Tipo_Operacion NVARCHAR(20) NULL,
    Cantidad_Registros INT NULL,
    Usuario NVARCHAR(100) NULL,
    IP_Origen NVARCHAR(50) NULL,
    Detalles NVARCHAR(MAX) NULL,
    Fecha_Operacion DATETIME NOT NULL CONSTRAINT DF_Auditoria_Fecha DEFAULT GETDATE()
);
GO

CREATE INDEX IX_CPU_Host ON Inventario.CPU(Host);
CREATE INDEX IX_CPU_Empresa ON Inventario.CPU(Empresa);
CREATE INDEX IX_CPU_Edificio ON Inventario.CPU(Edificio);
CREATE INDEX IX_CPU_Software_Host ON Inventario.CPU_Software(Host_CPU);
CREATE INDEX IX_CPU_Perifericos_Host ON Inventario.CPU_Perifericos(Host_CPU);
CREATE INDEX IX_CPU_Perifericos_Tipo ON Inventario.CPU_Perifericos(Tipo);
CREATE INDEX IX_CPU_Relaciones_CPU ON Inventario.CPU_Relaciones(Codigo_Barras_CPU);
CREATE INDEX IX_Otros_Tipo ON Inventario.Otros_Equipos(Tipo);
CREATE INDEX IX_Otros_Empresa ON Inventario.Otros_Equipos(Empresa);
CREATE INDEX IX_Otros_Edificio ON Inventario.Otros_Equipos(Edificio);
GO

PRINT 'Creando vistas...';
GO

CREATE VIEW Inventario.vw_CPU_Completo
AS
SELECT
    c.ID,
    c.Host,
    c.No_Serie,
    c.Empresa,
    c.Edificio,
    c.Area,
    c.Estado,
    c.Marca,
    c.Modelo,
    c.Procesador,
    c.RAM,
    c.Capacidad_Disco,
    c.Tipo_Disco,
    c.Observaciones,
    c.Codigo_QR,
    c.Codigo_Barras_CPU,
    c.Timestamp,
    s.SO,
    s.Office,
    s.Antivirus,
    s.Lector_PDF,
    s.ERP,
    s.Otro_1,
    s.Otro_2,
    s.Otro_3
FROM Inventario.CPU c
LEFT JOIN Inventario.CPU_Software s
    ON s.Host_CPU = c.Host;
GO

CREATE VIEW Inventario.vw_Otros_Equipos
AS
SELECT
    ID,
    Tipo,
    Nombre,
    No_Serie,
    Marca,
    Modelo,
    Empresa,
    Edificio,
    Area,
    Ubicacion_En_Edificio,
    Estado,
    Tipo_Sensor,
    Resolucion_Pantalla,
    Sistema_Operativo,
    Observaciones,
    Codigo_Barras,
    Codigo_ID,
    Timestamp
FROM Inventario.Otros_Equipos;
GO

PRINT '✅ Estructura Inventario creada correctamente.';
GO