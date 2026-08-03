-- Estructura de la tabla que aloja las empresas de "Empresas granel.xlsx"
-- (Granel) y "Empresas envasado.xlsb" (Envasado), con su CIF, código de
-- empresa, nombre y si están activas o no.
--
-- Esta tabla también se crea sola desde cargar_empresas.py la primera vez
-- que se ejecuta, así que ejecutar este script es opcional.

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EmpresasClasificadas')
CREATE TABLE EmpresasClasificadas (
    Id                  INT IDENTITY(1,1) PRIMARY KEY,
    [Clasificacion]     NVARCHAR(20) NOT NULL,   -- 'Granel' o 'Envasado'
    [CIF]               NVARCHAR(30) NOT NULL,
    [CodigoEmpresa]     NVARCHAR(20) NOT NULL,
    [NombreEmpresa]     NVARCHAR(200) NOT NULL,
    [Activa]            BIT NOT NULL DEFAULT 0,
    FechaCarga          DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT UQ_EmpresasClasificadas_CIF_Clasificacion UNIQUE (CIF, Clasificacion)
);
