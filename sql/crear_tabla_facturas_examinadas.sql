-- Estructura de la tabla que aloja las facturas "examinada" (extracción
-- correcta, ya sea automática o corregida manualmente).
--
-- Esta tabla también se crea sola desde la aplicación (ver
-- sql_historial.py) la primera vez que se guarda una factura, así que
-- ejecutar este script es opcional: solo hace falta si el DBA prefiere
-- crearla él mismo de antemano, con los permisos que corresponda.

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FacturasExaminadas')
CREATE TABLE FacturasExaminadas (
    Id                  INT IDENTITY(1,1) PRIMARY KEY,
    [Archivo]           NVARCHAR(255) NULL,
    [BaseImp]           NVARCHAR(255) NULL,
    [BaseIRPF]          NVARCHAR(255) NULL,
    [Buyer]             NVARCHAR(255) NULL,
    [Empresa]           NVARCHAR(255) NULL,
    [FEscaneo]          NVARCHAR(255) NULL,
    [FFactura]          NVARCHAR(255) NULL,
    [FOperacion]        NVARCHAR(255) NULL,
    [ImporIVA]          NVARCHAR(255) NULL,
    [Moneda]            NVARCHAR(255) NULL,
    [NombreProveedor]   NVARCHAR(255) NULL,
    [NumeroFactura]     NVARCHAR(255) NULL,
    [PedidoCliente]     NVARCHAR(1000) NULL,  -- puede traer varios pedidos concatenados con ";"
    [Proveedor]         NVARCHAR(255) NULL,
    [TipoIVA]           NVARCHAR(255) NULL,
    [TipoIVA2]          NVARCHAR(255) NULL,
    [TipoIVA3]          NVARCHAR(255) NULL,
    [TotalFact]         NVARCHAR(255) NULL,
    Origen              NVARCHAR(20) NULL,           -- 'auto' o 'manual'
    FechaInsercion      DATETIME NOT NULL DEFAULT GETDATE(),
    Definitiva          BIT NOT NULL DEFAULT 0,       -- 1 = marcada como 100% revisada/definitiva
    UsuarioDefinitiva   NVARCHAR(100) NULL,           -- usuario que la marcó/desmarcó
    FechaDefinitiva     DATETIME NULL,                -- fecha del último marcado/desmarcado
    Duplicado           BIT NOT NULL DEFAULT 0,        -- 1 = coincide con otra en NumeroFactura+Proveedor+Buyer
    CONSTRAINT UQ_FacturasExaminadas_Archivo UNIQUE (Archivo)
);
