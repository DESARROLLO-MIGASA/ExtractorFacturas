"""
Tests de auth.py (identidad/roles vía Windows Authentication + grupos de
AD). Puramente en memoria: no arrancan un servidor, no llaman a ningún
endpoint, no tocan la base de datos y no consultan Active Directory de
verdad (se simula _es_miembro_de_grupo) -así que se pueden ejecutar en
cualquier momento sin ningún efecto secundario.

Uso: python test_auth.py
"""

import unittest
from unittest.mock import patch

from fastapi import HTTPException

import auth


class TestNormalizarUsuario(unittest.TestCase):

    def test_formato_dominio_barra(self):
        self.assertEqual(auth._normalizar_usuario("DOMINT\\escaneria"), "escaneria")

    def test_formato_upn(self):
        self.assertEqual(auth._normalizar_usuario("escaneria@domint.local"), "escaneria")

    def test_sin_dominio(self):
        self.assertEqual(auth._normalizar_usuario("escaneria"), "escaneria")

    def test_pasa_a_minusculas(self):
        self.assertEqual(auth._normalizar_usuario("DOMINT\\EscanerIA"), "escaneria")

    def test_quita_espacios(self):
        self.assertEqual(auth._normalizar_usuario("  DOMINT\\escaneria  "), "escaneria")


def _grupo_falso(pertenece_a):
    """Sustituto de auth._es_miembro_de_grupo: `pertenece_a` es el conjunto
    de nombres de grupo (los valores de GRUPO_ADMINS/GRUPO_USUARIOS que se
    usen en cada test) de los que el usuario del test es miembro."""
    def fake(usuario, grupo):
        return grupo in pertenece_a
    return fake


class TestGetCurrentUser(unittest.TestCase):

    def test_sin_cabecera_lanza_401(self):
        with self.assertRaises(HTTPException) as ctx:
            auth.get_current_user(x_forwarded_user=None)
        self.assertEqual(ctx.exception.status_code, 401)

    def test_cabecera_vacia_lanza_401(self):
        with self.assertRaises(HTTPException) as ctx:
            auth.get_current_user(x_forwarded_user="   ")
        self.assertEqual(ctx.exception.status_code, 401)

    @patch.object(auth, "GRUPO_ADMINS", "GrupoAdmins")
    @patch.object(auth, "GRUPO_USUARIOS", "GrupoUsuarios")
    def test_miembro_del_grupo_admins_es_admin(self):
        with patch.object(auth, "_es_miembro_de_grupo", side_effect=_grupo_falso({"GrupoAdmins"})):
            usuario = auth.get_current_user(x_forwarded_user="DOMINT\\jperez")
        self.assertEqual(usuario["role"], "admin")
        self.assertEqual(usuario["username"], "jperez")
        self.assertEqual(usuario["usuario_dominio"], "DOMINT\\jperez")

    @patch.object(auth, "GRUPO_ADMINS", "GrupoAdmins")
    @patch.object(auth, "GRUPO_USUARIOS", "GrupoUsuarios")
    def test_miembro_del_grupo_usuarios_es_usuario(self):
        with patch.object(auth, "_es_miembro_de_grupo", side_effect=_grupo_falso({"GrupoUsuarios"})):
            usuario = auth.get_current_user(x_forwarded_user="DOMINT\\mgallego")
        self.assertEqual(usuario["role"], "usuario")

    @patch.object(auth, "GRUPO_ADMINS", "GrupoAdmins")
    @patch.object(auth, "GRUPO_USUARIOS", "GrupoUsuarios")
    def test_sin_pertenecer_a_ningun_grupo_lanza_403(self):
        with patch.object(auth, "_es_miembro_de_grupo", side_effect=_grupo_falso(set())):
            with self.assertRaises(HTTPException) as ctx:
                auth.get_current_user(x_forwarded_user="DOMINT\\forastero")
        self.assertEqual(ctx.exception.status_code, 403)

    @patch.object(auth, "GRUPO_ADMINS", "GrupoAdmins")
    @patch.object(auth, "GRUPO_USUARIOS", "GrupoUsuarios")
    def test_admin_tiene_prioridad_si_esta_en_los_dos_grupos(self):
        with patch.object(auth, "_es_miembro_de_grupo", side_effect=_grupo_falso({"GrupoAdmins", "GrupoUsuarios"})):
            usuario = auth.get_current_user(x_forwarded_user="DOMINT\\jperez")
        self.assertEqual(usuario["role"], "admin")


class TestEsMiembroDeGrupo(unittest.TestCase):

    def test_grupo_vacio_devuelve_false_sin_consultar_ad(self):
        # Si AD_GRUPO_ADMINS/AD_GRUPO_USUARIOS no están configurados en
        # .env, no debe intentar ninguna consulta ADSI (fallaría igual,
        # pero mejor no depender de eso): grupo="" -> False directo.
        self.assertFalse(auth._es_miembro_de_grupo("cualquiera", ""))


class TestRequerirAdmin(unittest.TestCase):

    def test_admin_pasa(self):
        usuario = {"username": "jperez", "usuario_dominio": "DOMINT\\jperez", "role": "admin"}
        self.assertEqual(auth.requerir_admin(usuario=usuario), usuario)

    def test_no_admin_lanza_403(self):
        usuario = {"username": "mgallego", "usuario_dominio": "DOMINT\\mgallego", "role": "usuario"}
        with self.assertRaises(HTTPException) as ctx:
            auth.requerir_admin(usuario=usuario)
        self.assertEqual(ctx.exception.status_code, 403)


if __name__ == "__main__":
    unittest.main()
