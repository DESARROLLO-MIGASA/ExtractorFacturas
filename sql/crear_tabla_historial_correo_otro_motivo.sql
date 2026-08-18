-- Histórico de envíos de correo del motivo "otros" vía Power Automate:
-- destinatario, asunto, motivo, cuerpo redactado, usuario, fecha y estado
-- (enviado/error) de cada intento. Permite en el futuro consultar desde
-- EscanerIA qué correos se han mandado sobre una factura.
--
-- Esta tabla también se intenta crear sola desde la aplicación (ver
-- sql_historial.py, _crear_tabla_historial_correo_otro_motivo_si_no_existe),
-- pero si el login de SQL no tiene permiso CREATE TABLE en la base de datos
-- (como ha pasado en producción con otras tablas), el intento automático
-- falla en silencio en cada llamada, así que un DBA debe ejecutar este
-- script una vez a mano.

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'HistorialCorreoOtroMotivo')
CREATE TABLE HistorialCorreoOtroMotivo (
    Id            INT IDENTITY(1,1) PRIMARY KEY,
    [Fecha]       DATETIME NOT NULL DEFAULT GETDATE(),
    [Archivo]     NVARCHAR(255) NOT NULL,
    [Destinatario] NVARCHAR(255) NULL,
    [Asunto]      NVARCHAR(500) NULL,
    [Motivo]      NVARCHAR(200) NOT NULL,
    [Cuerpo]      NVARCHAR(MAX) NOT NULL,
    [Usuario]     NVARCHAR(100) NOT NULL,
    [Estado]      NVARCHAR(20) NOT NULL,
    [MensajeError] NVARCHAR(1000) NULL
);
