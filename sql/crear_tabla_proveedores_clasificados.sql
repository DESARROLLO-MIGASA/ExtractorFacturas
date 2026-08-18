-- Estructura de la tabla que aloja los proveedores de ProveedoresBC.xls
-- (Granel) y ProveedoresBCEnvasado.xls (Envasado), con su CIF, nombre y
-- si están bloqueados o no.
--
-- Esta tabla también se crea sola desde cargar_proveedores.py la primera
-- vez que se ejecuta, así que ejecutar este script es opcional.
--
-- Si la tabla ya existía de antes (sin Direccion/Poblacion), cargar_proveedores.py
-- intenta añadir esas dos columnas solo, pero el login de SQL que usa la app
-- puede no tener permiso de ALTER TABLE (error 1088/"no existe o no tiene
-- permisos" aunque la tabla sí exista). En ese caso, ejecutar este script una
-- vez con un login con permiso de DDL deja el esquema listo; luego
-- cargar_proveedores.py ya puede rellenar los datos con el login normal.

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ProveedoresClasificados')
CREATE TABLE ProveedoresClasificados (
    Id                  INT IDENTITY(1,1) PRIMARY KEY,
    [Clasificacion]     NVARCHAR(20) NOT NULL,   -- 'Granel' o 'Envasado'
    [CIF]               NVARCHAR(30) NOT NULL,
    [NombreEmpresa]     NVARCHAR(200) NOT NULL,
    [Direccion]         NVARCHAR(200) NULL,
    [Poblacion]         NVARCHAR(100) NULL,
    [CodigoPostal]      NVARCHAR(10) NULL,
    [Bloqueado]         BIT NOT NULL DEFAULT 0,
    FechaCarga          DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT UQ_ProveedoresClasificados_CIF_Clasificacion UNIQUE (CIF, Clasificacion)
);

IF COL_LENGTH('ProveedoresClasificados', 'Direccion') IS NULL
ALTER TABLE ProveedoresClasificados ADD Direccion NVARCHAR(200) NULL;

IF COL_LENGTH('ProveedoresClasificados', 'Poblacion') IS NULL
ALTER TABLE ProveedoresClasificados ADD Poblacion NVARCHAR(100) NULL;

IF COL_LENGTH('ProveedoresClasificados', 'CodigoPostal') IS NULL
ALTER TABLE ProveedoresClasificados ADD CodigoPostal NVARCHAR(10) NULL;
