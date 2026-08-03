-- Estructura de la tabla de auditoría: quién hizo cada acción (editar,
-- confirmar, descartar, enviar correo, etc.), cuándo y sobre qué factura.
--
-- Esta tabla también se intenta crear sola desde la aplicación (ver
-- sql_historial.py, _crear_tabla_auditoria_si_no_existe), pero eso exige que
-- el login de SQL tenga permiso CREATE TABLE en la base de datos. Si ese
-- permiso no está concedido (como ha pasado en producción con otras tablas:
-- "Se ha denegado el permiso CREATE TABLE en la base de datos 'facturas'"),
-- el intento automático falla en SILENCIO en cada llamada -no solo la
-- primera vez-, así que TODO el registro de auditoría fallaría siempre
-- hasta que un DBA ejecute este script una vez a mano.

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'AuditLog')
CREATE TABLE AuditLog (
    Id          INT IDENTITY(1,1) PRIMARY KEY,
    [Fecha]     DATETIME NOT NULL DEFAULT GETDATE(),
    [Usuario]   NVARCHAR(100) NOT NULL,
    [Accion]    NVARCHAR(50) NOT NULL,
    [Archivo]   NVARCHAR(255) NULL,
    [Detalle]   NVARCHAR(500) NULL
);
