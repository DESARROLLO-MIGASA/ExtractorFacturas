-- Ejecutar UNA sola vez en la base de datos "facturas" (YBR-VM-SCANFATC),
-- con una cuenta que tenga permiso CREATE TABLE (db_owner o similar).
-- Después de esto, la app (con su cuenta normal, sin permisos de DDL)
-- ya puede leer/escribir en esta tabla sin problema, igual que hace hoy
-- con FacturasExaminadas / FacturasRevisadas / ReservasFacturas.

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ColaRevision')
CREATE TABLE ColaRevision (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    [Archivo] NVARCHAR(255) NULL,
    [NumeroFactura] NVARCHAR(255) NULL,
    [Buyer] NVARCHAR(255) NULL,
    [Empresa] NVARCHAR(255) NULL,
    [Proveedor] NVARCHAR(255) NULL,
    [NombreProveedor] NVARCHAR(255) NULL,
    [NumerosAlbaran] NVARCHAR(MAX) NULL,
    [PedidoCliente] NVARCHAR(MAX) NULL,
    [BaseImp] NVARCHAR(255) NULL,
    [BaseIRPF] NVARCHAR(255) NULL,
    [TipoIVA] NVARCHAR(255) NULL,
    [TipoIVA2] NVARCHAR(255) NULL,
    [TipoIVA3] NVARCHAR(255) NULL,
    [ImporIVA] NVARCHAR(255) NULL,
    [TotalFact] NVARCHAR(255) NULL,
    [Moneda] NVARCHAR(255) NULL,
    [FFactura] NVARCHAR(255) NULL,
    [FOperacion] NVARCHAR(255) NULL,
    [FEscaneo] NVARCHAR(255) NULL,
    Cola NVARCHAR(20) NOT NULL,
    FechaActualizacion DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT UQ_ColaRevision_Archivo UNIQUE (Archivo)
);
