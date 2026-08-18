"""
Tests del envío de correo "otro motivo" vía Power Automate (logic.py:
_texto_a_html_seguro, enviar_correo_power_automate, solicitar_correo_otro_motivo_pa).

Puramente en memoria, igual que test_auth.py: no llaman a Power Automate real
(se mockea httpx.post), no tocan SQL Server/SQLite (se mockea
registrar_historial_correo_otro_motivo_sql) y no mueven PDFs reales del flujo
de facturas (se mockea solicitar_envio_correo). Solo se usa un fichero
temporal como "PDF" para poder leerlo con open()/base64 de verdad.

Uso: python test_correo_otro_motivo.py
"""

import os
import tempfile
import unittest
from unittest.mock import patch

import httpx

import logic


URL_FLUJO_PRUEBA = "https://poweratomate.example.invalid/flujo-prueba"


class TestTextoAHtmlSeguro(unittest.TestCase):

    def test_escapa_html_y_convierte_parrafos(self):
        resultado = logic._texto_a_html_seguro(
            "Hola <b>Juan</b>,\n\nGracias por tu factura.\nUn saludo."
        )
        self.assertNotIn("<b>", resultado)
        self.assertIn("&lt;b&gt;", resultado)
        self.assertEqual(resultado.count("<p>"), 2)
        self.assertIn("<br>", resultado)

    def test_cadena_vacia_o_none(self):
        self.assertEqual(logic._texto_a_html_seguro(""), "")
        self.assertEqual(logic._texto_a_html_seguro(None), "")
        self.assertEqual(logic._texto_a_html_seguro("   "), "")


class TestEnviarCorreoPowerAutomate(unittest.TestCase):

    def setUp(self):
        fd, self.ruta_pdf = tempfile.mkstemp(suffix=".pdf")
        with os.fdopen(fd, "wb") as f:
            f.write(b"%PDF-1.4 contenido de prueba")

    def tearDown(self):
        os.remove(self.ruta_pdf)

    def _payload_base(self):
        return dict(
            destinatario="proveedor@ejemplo.com",
            asunto="Factura F001.pdf",
            cuerpo_html="<p>hola</p>",
            motivo="Falta el CIF",
            nombre_archivo="F001.pdf",
            ruta_pdf=self.ruta_pdf,
            usuario="jperez",
            archivo="F001.pdf",
        )

    @patch.dict(os.environ, {"POWER_AUTOMATE_CORREO_URL": ""})
    def test_no_configurado_lanza_runtimeerror(self):
        with self.assertRaises(RuntimeError):
            logic.enviar_correo_power_automate(**self._payload_base())

    @patch.dict(os.environ, {"POWER_AUTOMATE_CORREO_URL": URL_FLUJO_PRUEBA})
    @patch("logic.httpx.post")
    def test_envio_correcto(self, mock_post):
        mock_post.return_value = httpx.Response(200, request=httpx.Request("POST", URL_FLUJO_PRUEBA))

        resultado = logic.enviar_correo_power_automate(**self._payload_base())

        self.assertTrue(resultado)
        _, kwargs = mock_post.call_args
        self.assertEqual(kwargs["json"]["destinatario"], "proveedor@ejemplo.com")
        self.assertEqual(kwargs["json"]["archivo"], "F001.pdf")
        self.assertIn("pdf_base64", kwargs["json"])
        self.assertTrue(len(kwargs["json"]["pdf_base64"]) > 0)

    @patch.dict(os.environ, {"POWER_AUTOMATE_CORREO_URL": URL_FLUJO_PRUEBA})
    @patch("logic.httpx.post", side_effect=httpx.TimeoutException("timeout"))
    def test_timeout_lanza_runtimeerror(self, mock_post):
        with self.assertRaises(RuntimeError):
            logic.enviar_correo_power_automate(**self._payload_base())

    @patch.dict(os.environ, {"POWER_AUTOMATE_CORREO_URL": URL_FLUJO_PRUEBA})
    @patch("logic.httpx.post")
    def test_respuesta_error_lanza_runtimeerror(self, mock_post):
        mock_post.return_value = httpx.Response(
            500, request=httpx.Request("POST", URL_FLUJO_PRUEBA), text="boom"
        )
        with self.assertRaises(RuntimeError):
            logic.enviar_correo_power_automate(**self._payload_base())

    @patch.dict(os.environ, {"POWER_AUTOMATE_CORREO_URL": URL_FLUJO_PRUEBA})
    @patch("logic.httpx.post", side_effect=httpx.ConnectError("no hay red"))
    def test_error_conexion_lanza_runtimeerror(self, mock_post):
        with self.assertRaises(RuntimeError):
            logic.enviar_correo_power_automate(**self._payload_base())


class TestSolicitarCorreoOtroMotivoPA(unittest.TestCase):

    def setUp(self):
        fd, self.ruta_pdf = tempfile.mkstemp(suffix=".pdf")
        with os.fdopen(fd, "wb") as f:
            f.write(b"%PDF-1.4 contenido de prueba")
        self.datos_ok = {
            "email": "proveedor@ejemplo.com",
            "nombre_adjunto": "F001.pdf",
            "asunto": "Factura F001.pdf",
        }

    def tearDown(self):
        os.remove(self.ruta_pdf)

    def test_motivo_vacio_lanza_valueerror(self):
        with self.assertRaises(ValueError):
            logic.solicitar_correo_otro_motivo_pa("F001.pdf", "   ", "cuerpo", "jperez")

    def test_cuerpo_vacio_lanza_valueerror(self):
        with patch.object(logic, "datos_correo_outlook", return_value=self.datos_ok):
            with self.assertRaises(ValueError):
                logic.solicitar_correo_otro_motivo_pa("F001.pdf", "motivo", "   ", "jperez")

    def test_factura_sin_email_lanza_valueerror(self):
        datos_sin_email = dict(self.datos_ok, email="")
        with patch.object(logic, "datos_correo_outlook", return_value=datos_sin_email):
            with self.assertRaises(ValueError):
                logic.solicitar_correo_otro_motivo_pa("F001.pdf", "motivo", "cuerpo", "jperez")

    def test_pdf_inexistente_lanza_filenotfounderror(self):
        # datos_correo_outlook ya lanzaría FileNotFoundError en un caso real
        # (llama a buscar_pdf_por_nombre internamente), pero aquí se simula
        # el hueco -PDF que existía al pintar el modal pero desaparece justo
        # antes de enviar- para probar la segunda comprobación explícita.
        with patch.object(logic, "datos_correo_outlook", return_value=self.datos_ok), \
             patch.object(logic, "buscar_pdf_por_nombre", return_value=None):
            with self.assertRaises(FileNotFoundError):
                logic.solicitar_correo_otro_motivo_pa("F001.pdf", "motivo", "cuerpo", "jperez")

    def test_pdf_inexistente_desde_datos_correo_outlook(self):
        with patch.object(logic, "datos_correo_outlook", side_effect=FileNotFoundError("F001.pdf")):
            with self.assertRaises(FileNotFoundError):
                logic.solicitar_correo_otro_motivo_pa("F001.pdf", "motivo", "cuerpo", "jperez")

    def test_envio_correcto_registra_historial_y_archiva_pdf(self):
        with patch.object(logic, "datos_correo_outlook", return_value=self.datos_ok), \
             patch.object(logic, "buscar_pdf_por_nombre", return_value=self.ruta_pdf), \
             patch.object(logic, "_numero_factura_de_archivo", return_value="26/1224"), \
             patch.object(logic, "enviar_correo_power_automate", return_value=True) as mock_enviar, \
             patch.object(logic, "registrar_historial_correo_otro_motivo_sql") as mock_historial, \
             patch.object(logic, "solicitar_envio_correo") as mock_solicitar:

            resultado = logic.solicitar_correo_otro_motivo_pa(
                "F001.pdf", "Falta el CIF", "Hola,\n\nAdjunto factura.", "jperez"
            )

        self.assertTrue(resultado)
        mock_enviar.assert_called_once()
        self.assertEqual(mock_enviar.call_args[1]["asunto"], "Factura 26/1224 — Falta el CIF")
        mock_historial.assert_called_once()
        self.assertEqual(mock_historial.call_args[0][2], "Factura 26/1224 — Falta el CIF")
        self.assertEqual(mock_historial.call_args[0][6], "enviado")
        mock_solicitar.assert_called_once_with("F001.pdf", "otros", "jperez", motivo_otro="Falta el CIF")

    def test_envio_correcto_sin_numero_factura_usa_nombre_adjunto(self):
        # Si no hay número de factura conocido (ni en FacturasExaminadas ni
        # en ColaRevision), el asunto cae de vuelta al nombre del PDF en vez
        # de romper el envío.
        with patch.object(logic, "datos_correo_outlook", return_value=self.datos_ok), \
             patch.object(logic, "buscar_pdf_por_nombre", return_value=self.ruta_pdf), \
             patch.object(logic, "_numero_factura_de_archivo", return_value=None), \
             patch.object(logic, "enviar_correo_power_automate", return_value=True) as mock_enviar, \
             patch.object(logic, "registrar_historial_correo_otro_motivo_sql"), \
             patch.object(logic, "solicitar_envio_correo"):

            logic.solicitar_correo_otro_motivo_pa("F001.pdf", "Falta el CIF", "cuerpo", "jperez")

        self.assertEqual(mock_enviar.call_args[1]["asunto"], "Factura F001.pdf — Falta el CIF")

    def test_error_power_automate_registra_historial_error_y_no_archiva(self):
        with patch.object(logic, "datos_correo_outlook", return_value=self.datos_ok), \
             patch.object(logic, "buscar_pdf_por_nombre", return_value=self.ruta_pdf), \
             patch.object(logic, "_numero_factura_de_archivo", return_value="26/1224"), \
             patch.object(logic, "enviar_correo_power_automate",
                          side_effect=RuntimeError("Power Automate rechazó el envío del correo.")), \
             patch.object(logic, "registrar_historial_correo_otro_motivo_sql") as mock_historial, \
             patch.object(logic, "solicitar_envio_correo") as mock_solicitar:

            with self.assertRaises(RuntimeError):
                logic.solicitar_correo_otro_motivo_pa("F001.pdf", "Falta el CIF", "cuerpo", "jperez")

        mock_historial.assert_called_once()
        self.assertEqual(mock_historial.call_args[0][2], "Factura 26/1224 — Falta el CIF")
        self.assertEqual(mock_historial.call_args[0][6], "error")
        self.assertIn("mensaje_error", mock_historial.call_args[1])
        mock_solicitar.assert_not_called()


class TestNumeroFacturaDeArchivo(unittest.TestCase):

    def test_usa_facturas_examinadas_si_existe(self):
        with patch.object(logic, "obtener_factura_examinada_sql", return_value={"NumeroFactura": "26/1224"}):
            self.assertEqual(logic._numero_factura_de_archivo("F001.pdf"), "26/1224")

    def test_cae_a_cola_revision_si_no_esta_examinada(self):
        idx_archivo = logic.COLUMNAS.index("Archivo")
        idx_numero = logic.COLUMNAS.index("NumeroFactura")
        fila = [None] * len(logic.COLUMNAS)
        fila[idx_archivo] = "F001.pdf"
        fila[idx_numero] = "A-99"

        with patch.object(logic, "obtener_factura_examinada_sql", return_value=None), \
             patch.object(logic, "listar_cola_revision_sql", side_effect=lambda cola: [fila] if cola == "incidencia" else []):
            self.assertEqual(logic._numero_factura_de_archivo("F001.pdf"), "A-99")

    def test_devuelve_none_si_no_se_encuentra_en_ningun_sitio(self):
        with patch.object(logic, "obtener_factura_examinada_sql", return_value=None), \
             patch.object(logic, "listar_cola_revision_sql", return_value=[]):
            self.assertIsNone(logic._numero_factura_de_archivo("F001.pdf"))


if __name__ == "__main__":
    unittest.main()
