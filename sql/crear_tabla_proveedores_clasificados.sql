-- Estructura de la tabla que aloja los proveedores de ProveedoresBC.xls
-- (Granel) y ProveedoresBCEnvasado.xls (Envasado), con su CIF, nombre y
-- si están bloqueados o no.
--
-- Esta tabla también se crea sola desde cargar_proveedores.py la primera
-- vez que se ejecuta, así que ejecutar este script es opcional.

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ProveedoresClasificados')
CREATE TABLE ProveedoresClasificados (
    Id                  INT IDENTITY(1,1) PRIMARY KEY,
    [Clasificacion]     NVARCHAR(20) NOT NULL,   -- 'Granel' o 'Envasado'
    [CIF]               NVARCHAR(30) NOT NULL,
    [NombreEmpresa]     NVARCHAR(200) NOT NULL,
    [Bloqueado]         BIT NOT NULL DEFAULT 0,
    FechaCarga          DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT UQ_ProveedoresClasificados_CIF_Clasificacion UNIQUE (CIF, Clasificacion)
);
