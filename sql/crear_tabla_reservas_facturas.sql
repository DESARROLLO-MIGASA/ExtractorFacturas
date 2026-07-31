-- Estructura de la tabla que guarda qué factura tiene reservada cada
-- persona ahora mismo (para que no se pisen dos revisores a la vez).
--
-- Esta tabla también se intenta crear sola desde la aplicación (ver
-- sql_historial.py, _crear_tabla_reservas_si_no_existe), pero eso exige que
-- el login de SQL tenga permiso CREATE TABLE en la base de datos. Si ese
-- permiso no está concedido (como ha pasado en producción: "Se ha denegado
-- el permiso CREATE TABLE en la base de datos 'facturas'"), el intento
-- automático falla en SILENCIO en cada llamada -no solo la primera vez-,
-- así que TODAS las reservas fallan siempre y la función de reservas queda
-- inservible hasta que un DBA ejecute este script una vez a mano.

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ReservasFacturas')
CREATE TABLE ReservasFacturas (
    Id                  INT IDENTITY(1,1) PRIMARY KEY,
    [Archivo]           NVARCHAR(255) NOT NULL,
    [Usuario]           NVARCHAR(100) NOT NULL,
    [FechaReserva]      DATETIME NOT NULL,
    [UltimaActividad]   DATETIME NOT NULL,
    CONSTRAINT UQ_ReservasFacturas_Archivo UNIQUE (Archivo)
);
