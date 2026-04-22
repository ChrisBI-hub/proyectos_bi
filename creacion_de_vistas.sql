-- ============================================================
-- SCRIPTS SQL PARA CREAR TABLAS v2.1
-- Ejecutar en SQL Server ANTES de usar sync_bd.py
-- v2.1: Columnas de texto ampliadas para rutas de archivo
-- ============================================================

USE BI;
GO

-- ============================================================
-- 1. TABLA PRINCIPAL: CPU
-- ============================================================
IF OBJECT_ID('Inventario.CPU', 'U') IS NULL
BEGIN
    CREATE TABLE Inventario.CPU (
        ID INT PRIMARY KEY IDENTITY(1,1),
        Host NVARCHAR(100) NOT NULL UNIQUE,
        No_Serie NVARCHAR(100) NOT NULL,
        Empresa NVARCHAR(200),
        Edificio NVARCHAR(100),
        Area NVARCHAR(200),
        Estado NVARCHAR(50),
        Marca NVARCHAR(100),
        Modelo NVARCHAR(100),
        Procesador NVARCHAR(150),
        RAM NVARCHAR(50),
        Capacidad_Disco NVARCHAR(50),
        Tipo_Disco NVARCHAR(20),
        Observaciones NVARCHAR(MAX),
        Codigo_QR NVARCHAR(500),
        Codigo_Barras_CPU NVARCHAR(255) UNIQUE,
        Timestamp DATETIME DEFAULT GETDATE(),
        Fecha_Creacion DATETIME DEFAULT GETDATE(),
        Fecha_Modificacion DATETIME DEFAULT GETDATE()
    );

    CREATE INDEX idx_CPU_Host ON Inventario.CPU(Host);
    CREATE INDEX idx_CPU_Empresa ON Inventario.CPU(Empresa);
    CREATE INDEX idx_CPU_Edificio ON Inventario.CPU(Edificio);
    
    PRINT 'Tabla Inventario.CPU creada exitosamente.';
END
ELSE
    PRINT 'Tabla Inventario.CPU ya existe (se aplicarán ALTER si es necesario).';
GO

-- ============================================================
-- 2. TABLA DE SOFTWARE INSTALADO EN CPU
-- ============================================================
IF OBJECT_ID('Inventario.CPU_Software', 'U') IS NULL
BEGIN
    CREATE TABLE Inventario.CPU_Software (
        ID INT PRIMARY KEY IDENTITY(1,1),
        Host_CPU NVARCHAR(100) NOT NULL,
        SO NVARCHAR(100),
        Office NVARCHAR(100),
        Antivirus NVARCHAR(100),
        Lector_PDF NVARCHAR(100),
        ERP NVARCHAR(100),
        Otro_1 NVARCHAR(100),
        Otro_2 NVARCHAR(100),
        Otro_3 NVARCHAR(100),
        Timestamp DATETIME DEFAULT GETDATE(),
        Fecha_Creacion DATETIME DEFAULT GETDATE(),
        
        CONSTRAINT FK_CPU_Software_Host 
            FOREIGN KEY (Host_CPU) 
            REFERENCES Inventario.CPU(Host)
            ON DELETE CASCADE
            ON UPDATE CASCADE
    );

    CREATE INDEX idx_CPU_Software_Host ON Inventario.CPU_Software(Host_CPU);
    PRINT 'Tabla Inventario.CPU_Software creada exitosamente.';
END
ELSE
    PRINT 'Tabla Inventario.CPU_Software ya existe.';
GO

-- ============================================================
-- 3. TABLA DE PERIFÉRICOS ASOCIADOS A CPU
-- ============================================================
IF OBJECT_ID('Inventario.CPU_Perifericos', 'U') IS NULL
BEGIN
    CREATE TABLE Inventario.CPU_Perifericos (
        ID INT PRIMARY KEY IDENTITY(1,1),
        Host_CPU NVARCHAR(100) NOT NULL,
        Tipo NVARCHAR(50) NOT NULL,
        Modelo NVARCHAR(100),
        No_Serie NVARCHAR(100),
        Marca NVARCHAR(100),
        Estado NVARCHAR(50),
        Observaciones NVARCHAR(MAX),
        Codigo_Barras NVARCHAR(255),
        Codigo_ID NVARCHAR(100) UNIQUE,
        Timestamp DATETIME DEFAULT GETDATE(),
        Fecha_Creacion DATETIME DEFAULT GETDATE(),
        
        CONSTRAINT FK_CPU_Perifericos_Host 
            FOREIGN KEY (Host_CPU) 
            REFERENCES Inventario.CPU(Host)
            ON DELETE CASCADE
            ON UPDATE CASCADE
    );

    CREATE INDEX idx_CPU_Perifericos_Host ON Inventario.CPU_Perifericos(Host_CPU);
    CREATE INDEX idx_CPU_Perifericos_Tipo ON Inventario.CPU_Perifericos(Tipo);
    CREATE INDEX idx_CPU_Perifericos_Codigo ON Inventario.CPU_Perifericos(Codigo_ID);
    PRINT 'Tabla Inventario.CPU_Perifericos creada exitosamente.';
END
ELSE
    PRINT 'Tabla Inventario.CPU_Perifericos ya existe.';
GO

-- ============================================================
-- 4. TABLA DE RELACIONES (QR ↔ CÓDIGOS DE BARRAS)
-- ============================================================
IF OBJECT_ID('Inventario.CPU_Relaciones', 'U') IS NULL
BEGIN
    CREATE TABLE Inventario.CPU_Relaciones (
        ID INT PRIMARY KEY IDENTITY(1,1),
        Codigo_QR NVARCHAR(500),
        Codigo_Barras_CPU NVARCHAR(255),
        Tipo_Periferico NVARCHAR(50),
        Codigo_Barras_Periferico NVARCHAR(255),
        Timestamp DATETIME DEFAULT GETDATE()
    );

    CREATE INDEX idx_CPU_Relaciones_QR ON Inventario.CPU_Relaciones(Codigo_QR);
    CREATE INDEX idx_CPU_Relaciones_Barras_CPU ON Inventario.CPU_Relaciones(Codigo_Barras_CPU);
    PRINT 'Tabla Inventario.CPU_Relaciones creada exitosamente.';
END
ELSE
    PRINT 'Tabla Inventario.CPU_Relaciones ya existe.';
GO

-- ============================================================
-- 5. TABLA DE OTROS EQUIPAMIENTOS (Sillas, Mesas, etc.)
-- ============================================================
IF OBJECT_ID('Inventario.Otros_Equipos', 'U') IS NULL
BEGIN
    CREATE TABLE Inventario.Otros_Equipos (
        ID INT PRIMARY KEY IDENTITY(1,1),
        Tipo NVARCHAR(100) NOT NULL,
        Nombre NVARCHAR(500),
        No_Serie NVARCHAR(200),
        Marca NVARCHAR(200),
        Modelo NVARCHAR(200),
        Empresa NVARCHAR(200),
        Edificio NVARCHAR(100),
        Area NVARCHAR(200),
        Estado NVARCHAR(50),
        Observaciones NVARCHAR(MAX),
        Codigo_Barras NVARCHAR(255),
        Codigo_ID NVARCHAR(100) UNIQUE,
        Timestamp DATETIME DEFAULT GETDATE(),
        Fecha_Creacion DATETIME DEFAULT GETDATE(),
        Fecha_Modificacion DATETIME DEFAULT GETDATE()
    );

    CREATE INDEX idx_Otros_Tipo ON Inventario.Otros_Equipos(Tipo);
    CREATE INDEX idx_Otros_Empresa ON Inventario.Otros_Equipos(Empresa);
    CREATE INDEX idx_Otros_Edificio ON Inventario.Otros_Equipos(Edificio);
    CREATE INDEX idx_Otros_Codigo ON Inventario.Otros_Equipos(Codigo_ID);
    PRINT 'Tabla Inventario.Otros_Equipos creada exitosamente.';
END
ELSE
    PRINT 'Tabla Inventario.Otros_Equipos ya existe.';
GO

-- ============================================================
-- 6. TABLA DE AUDITORÍA
-- ============================================================
IF OBJECT_ID('Inventario.Auditoria', 'U') IS NULL
BEGIN
    CREATE TABLE Inventario.Auditoria (
        ID INT PRIMARY KEY IDENTITY(1,1),
        Tabla_Afectada NVARCHAR(100),
        Tipo_Operacion NVARCHAR(20),
        Cantidad_Registros INT,
        Usuario NVARCHAR(100),
        IP_Origen NVARCHAR(50),
        Detalles NVARCHAR(MAX),
        Fecha_Operacion DATETIME DEFAULT GETDATE()
    );

    CREATE INDEX idx_Auditoria_Tabla ON Inventario.Auditoria(Tabla_Afectada);
    CREATE INDEX idx_Auditoria_Fecha ON Inventario.Auditoria(Fecha_Operacion);
    PRINT 'Tabla Inventario.Auditoria creada exitosamente.';
END
ELSE
    PRINT 'Tabla Inventario.Auditoria ya existe.';
GO

-- ============================================================
-- ALTER TABLE — Ampliar columnas en tablas ya existentes
-- Ejecutar siempre: si la columna ya es del tamaño correcto
-- SQL Server lo aplica sin error ni pérdida de datos.
-- ============================================================
PRINT '';
PRINT '🔧 Aplicando correcciones de columnas (ALTER TABLE)...';
GO

-- Inventario.Otros_Equipos
ALTER TABLE Inventario.Otros_Equipos ALTER COLUMN Codigo_Barras NVARCHAR(255);
ALTER TABLE Inventario.Otros_Equipos ALTER COLUMN Codigo_ID     NVARCHAR(100);
ALTER TABLE Inventario.Otros_Equipos ALTER COLUMN Nombre        NVARCHAR(500);
ALTER TABLE Inventario.Otros_Equipos ALTER COLUMN No_Serie      NVARCHAR(200);
ALTER TABLE Inventario.Otros_Equipos ALTER COLUMN Marca         NVARCHAR(200);
ALTER TABLE Inventario.Otros_Equipos ALTER COLUMN Modelo        NVARCHAR(200);
ALTER TABLE Inventario.Otros_Equipos ALTER COLUMN Empresa       NVARCHAR(200);
ALTER TABLE Inventario.Otros_Equipos ALTER COLUMN Edificio      NVARCHAR(100);
ALTER TABLE Inventario.Otros_Equipos ALTER COLUMN Area          NVARCHAR(200);
PRINT '   ✅ Inventario.Otros_Equipos — columnas ampliadas.';
GO

-- Inventario.CPU_Perifericos
ALTER TABLE Inventario.CPU_Perifericos ALTER COLUMN Codigo_Barras NVARCHAR(255);
ALTER TABLE Inventario.CPU_Perifericos ALTER COLUMN Codigo_ID     NVARCHAR(100);
PRINT '   ✅ Inventario.CPU_Perifericos — columnas ampliadas.';
GO

-- Inventario.CPU
ALTER TABLE Inventario.CPU ALTER COLUMN Codigo_Barras_CPU NVARCHAR(255);
ALTER TABLE Inventario.CPU ALTER COLUMN Codigo_QR         NVARCHAR(500);
ALTER TABLE Inventario.CPU ALTER COLUMN Empresa           NVARCHAR(200);
ALTER TABLE Inventario.CPU ALTER COLUMN Area              NVARCHAR(200);
ALTER TABLE Inventario.CPU ALTER COLUMN Edificio          NVARCHAR(100);
PRINT '   ✅ Inventario.CPU — columnas ampliadas.';
GO

-- Inventario.CPU_Relaciones
ALTER TABLE Inventario.CPU_Relaciones ALTER COLUMN Codigo_QR                NVARCHAR(500);
ALTER TABLE Inventario.CPU_Relaciones ALTER COLUMN Codigo_Barras_CPU        NVARCHAR(255);
ALTER TABLE Inventario.CPU_Relaciones ALTER COLUMN Codigo_Barras_Periferico NVARCHAR(255);
PRINT '   ✅ Inventario.CPU_Relaciones — columnas ampliadas.';
GO

-- ============================================================
-- 7. VISTAS ÚTILES
-- ============================================================

IF OBJECT_ID('Inventario.vw_CPU_Completa', 'V') IS NOT NULL
    DROP VIEW Inventario.vw_CPU_Completa;
GO
CREATE VIEW Inventario.vw_CPU_Completa AS
SELECT 
    cpu.ID, cpu.Host, cpu.No_Serie, cpu.Empresa, cpu.Edificio,
    cpu.Area, cpu.Estado, cpu.Marca, cpu.Modelo, cpu.Procesador,
    cpu.RAM, cpu.Capacidad_Disco, cpu.Tipo_Disco,
    cpu.Codigo_QR, cpu.Codigo_Barras_CPU,
    sw.SO, sw.Office, sw.Antivirus,
    COUNT(DISTINCT per.ID) AS Total_Perifericos,
    cpu.Timestamp
FROM Inventario.CPU cpu
LEFT JOIN Inventario.CPU_Software sw ON cpu.Host = sw.Host_CPU
LEFT JOIN Inventario.CPU_Perifericos per ON cpu.Host = per.Host_CPU
GROUP BY 
    cpu.ID, cpu.Host, cpu.No_Serie, cpu.Empresa, cpu.Edificio,
    cpu.Area, cpu.Estado, cpu.Marca, cpu.Modelo, cpu.Procesador,
    cpu.RAM, cpu.Capacidad_Disco, cpu.Tipo_Disco,
    cpu.Codigo_QR, cpu.Codigo_Barras_CPU,
    sw.SO, sw.Office, sw.Antivirus, cpu.Timestamp;
GO
PRINT 'Vista Inventario.vw_CPU_Completa creada/actualizada.';
GO

IF OBJECT_ID('Inventario.vw_Perifericos_CPU', 'V') IS NOT NULL
    DROP VIEW Inventario.vw_Perifericos_CPU;
GO
CREATE VIEW Inventario.vw_Perifericos_CPU AS
SELECT 
    cpu.Host AS CPU_Host, cpu.Codigo_Barras_CPU,
    per.Tipo AS Tipo_Periferico, per.Modelo, per.Marca,
    per.Estado, per.Codigo_ID, per.Timestamp
FROM Inventario.CPU_Perifericos per
INNER JOIN Inventario.CPU cpu ON per.Host_CPU = cpu.Host;
GO
PRINT 'Vista Inventario.vw_Perifericos_CPU creada/actualizada.';
GO

IF OBJECT_ID('Inventario.vw_Resumen_Empresa', 'V') IS NOT NULL
    DROP VIEW Inventario.vw_Resumen_Empresa;
GO
CREATE VIEW Inventario.vw_Resumen_Empresa AS
SELECT 
    Empresa,
    COUNT(DISTINCT ID) AS Total_CPUs,
    COUNT(DISTINCT CASE WHEN Estado = 'Buen estado'            THEN ID END) AS Buen_Estado,
    COUNT(DISTINCT CASE WHEN Estado = 'Futuro mantenimiento'   THEN ID END) AS Futuro_Mantenimiento,
    COUNT(DISTINCT CASE WHEN Estado = 'Mal estado'             THEN ID END) AS Mal_Estado,
    COUNT(DISTINCT CASE WHEN Estado = 'Equipo nuevo'           THEN ID END) AS Equipo_Nuevo
FROM Inventario.CPU
GROUP BY Empresa;
GO
PRINT 'Vista Inventario.vw_Resumen_Empresa creada/actualizada.';
GO

-- ============================================================
-- 8. STORED PROCEDURES
-- ============================================================

IF OBJECT_ID('Inventario.sp_ObtenerCPU_Detalles', 'P') IS NOT NULL
    DROP PROCEDURE Inventario.sp_ObtenerCPU_Detalles;
GO
CREATE PROCEDURE Inventario.sp_ObtenerCPU_Detalles
    @Host NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT 'CPU'         AS Categoria, * FROM Inventario.CPU              WHERE Host     = @Host;
    SELECT 'Software'    AS Categoria, * FROM Inventario.CPU_Software     WHERE Host_CPU = @Host;
    SELECT 'Periféricos' AS Categoria, * FROM Inventario.CPU_Perifericos  WHERE Host_CPU = @Host;
END;
GO
PRINT 'SP sp_ObtenerCPU_Detalles creado/actualizado.';
GO

IF OBJECT_ID('Inventario.sp_Equipos_Por_Edificio', 'P') IS NOT NULL
    DROP PROCEDURE Inventario.sp_Equipos_Por_Edificio;
GO
CREATE PROCEDURE Inventario.sp_Equipos_Por_Edificio
    @Edificio NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT 'CPU' AS Tipo, Host AS Nombre, No_Serie, Marca, Modelo, Estado, Timestamp
    FROM Inventario.CPU WHERE Edificio = @Edificio
    UNION ALL
    SELECT Tipo, Nombre, No_Serie, Marca, Modelo, Estado, Timestamp
    FROM Inventario.Otros_Equipos WHERE Edificio = @Edificio
    ORDER BY Tipo, Nombre;
END;
GO
PRINT 'SP sp_Equipos_Por_Edificio creado/actualizado.';
GO

-- ============================================================
-- 9. VERIFICACIÓN FINAL
-- ============================================================
PRINT '';
PRINT '✅ VERIFICACIÓN DE TABLAS:';
SELECT name AS Tabla, CONVERT(VARCHAR(19), create_date, 121) AS Fecha_Creacion
FROM sys.objects WHERE type = 'U' AND schema_id = SCHEMA_ID('Inventario') ORDER BY name;

PRINT '';
PRINT '✅ VERIFICACIÓN DE VISTAS:';
SELECT name AS Vista, CONVERT(VARCHAR(19), create_date, 121) AS Fecha_Creacion
FROM sys.objects WHERE type = 'V' AND schema_id = SCHEMA_ID('Inventario') ORDER BY name;

PRINT '';
PRINT '═══════════════════════════════════════════════════════';
PRINT '✅ INSTALACIÓN v2.1 COMPLETADA EXITOSAMENTE';
PRINT '═══════════════════════════════════════════════════════';
GO