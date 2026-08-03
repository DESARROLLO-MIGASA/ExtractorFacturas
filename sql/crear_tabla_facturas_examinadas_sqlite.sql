-- Equivalente SQLite del script de SQL Server (crear_tabla_facturas_examinadas.sql).
-- Uso temporal mientras no hay acceso al SQL Server corporativo (ver sql_historial.py,
-- variable SQL_ENGINE en .env). Igual que en SQL Server, la tabla también se crea sola
-- desde la aplicación la primera vez que se guarda una factura, así que ejecutar este
-- script es opcional.

CREATE TABLE IF NOT EXISTS FacturasExaminadas (
    Id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    [Archivo]           TEXT,
    [BaseImp]           TEXT,
    [BaseIRPF]          TEXT,
    [Buyer]             TEXT,
    [Empresa]           TEXT,
    [FEscaneo]          TEXT,
    [FFactura]          TEXT,
    [FOperacion]        TEXT,
    [ImporIVA]          TEXT,
    [Moneda]            TEXT,
    [NombreProveedor]   TEXT,
    [NumeroFactura]     TEXT,
    [PedidoCliente]     TEXT,
    [Proveedor]         TEXT,
    [TipoIVA]           TEXT,
    [TipoIVA2]          TEXT,
    [TipoIVA3]          TEXT,
    [TotalFact]         TEXT,
    Origen              TEXT,                          -- 'auto' o 'manual'
    FechaInsercion      TEXT NOT NULL DEFAULT (datetime('now')),
    Definitiva          INTEGER DEFAULT 0,              -- 1 = marcada como 100% revisada/definitiva
    UsuarioDefinitiva   TEXT,                            -- usuario que la marcó/desmarcó
    FechaDefinitiva     TEXT,                            -- fecha del último marcado/desmarcado
    Duplicado           INTEGER DEFAULT 0,               -- 1 = coincide con otra en NumeroFactura+Proveedor+Buyer
    UNIQUE([Archivo])
);
