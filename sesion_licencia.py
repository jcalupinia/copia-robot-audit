"""Politica de sesion y licencia, separada de la interfaz.

Vive aparte de `aplicacion.py` por una razon concreta: es la logica que decide
si a alguien se le cierra la sesion, y enterrada dentro de un script de
Streamlit no habia forma de probarla. Los errores que tuvo -- cerrar sesion por
un timeout, pedir el codigo de licencia porque el servidor tardo en responder --
son justamente los que un test detecta en un segundo y un usuario descubre a la
mitad de una descarga.

Todas las funciones reciben `estado`, que en produccion es `st.session_state` y
en los tests es un diccionario comun. No importan Streamlit.

LA REGLA QUE ORDENA TODO EL MODULO
----------------------------------
Se distingue entre que el servidor CONTESTE que no y que NO SE PUEDA LLEGAR al
servidor:

- Rechazo explicito  -> es autoritativo: se cierra sesion o se pide activacion.
- Servidor sin responder -> no concluye nada: se sigue con lo que ya se sabia.

Nunca al reves. Un corte de internet no es una licencia vencida.
"""
from __future__ import annotations

import time
from typing import Callable, Optional

from licensing_client import RespuestaDelServidor, ServidorInalcanzable

# Cada cuanto se le vuelve a preguntar al servidor por una licencia ya validada.
REVALIDAR_HORAS_POR_DEFECTO = 12.0

# Cuanto sigue andando la app si el servidor no contesta. No es permiso para
# usarla sin licencia: es el margen para que un corte de red no deje a nadie
# afuera de su propio trabajo.
GRACIA_DIAS_POR_DEFECTO = 30.0

CLAVES_DE_SESION = (
    "auth_token",
    "refresh_token",
    "user_email",
    "license_validated",
    "license_last_check",
    "license_valid_until",
)


def _sin_efecto() -> None:
    return None


def renovar_token(estado, cliente, persistir: Callable[[], None] = _sin_efecto) -> str:
    """Canjea el refresh token por un access token nuevo.

    Devuelve "renovado", "sin_refresh", "sin_respuesta" o "rechazado". Son
    cuatro desenlaces distintos y hay que tratarlos distinto: solo "rechazado"
    y "sin_refresh" terminan la sesion.

    Es lo que hace que la sesion sobreviva al vencimiento del access token, que
    dura una hora. Sin esto, a los 60 minutos el usuario salia disparado sin
    haber tocado nada: el sintoma que se reportaba como "se cierra de la nada".
    """
    refresh = estado.get("refresh_token")
    if not refresh:
        return "sin_refresh"
    try:
        sesion = cliente.refresh(refresh)
    except ServidorInalcanzable:
        # No se pudo preguntar. No se concluye nada: se reintenta mas tarde.
        return "sin_respuesta"
    except RespuestaDelServidor:
        # El refresh token ya no vale: esta si es una sesion terminada.
        estado.pop("refresh_token", None)
        return "rechazado"
    estado["auth_token"] = sesion["access_token"]
    if sesion.get("refresh_token"):
        estado["refresh_token"] = sesion["refresh_token"]
    persistir()
    return "renovado"


def sesion_sigue_viva(estado, cliente, persistir: Callable[[], None] = _sin_efecto) -> bool:
    """Confirma la sesion, renovando el token si hace falta.

    Devuelve False SOLO cuando el servidor rechaza explicitamente y ademas no
    se pudo renovar. Si el servidor no contesta devuelve True: no poder
    preguntar no es motivo para echar a nadie.
    """
    if not estado.get("auth_token"):
        return False
    for intento in (1, 2):
        try:
            cliente.get_profile(estado["auth_token"])
            return True
        except ServidorInalcanzable:
            return True
        except RespuestaDelServidor as err:
            if not err.es_sesion_invalida:
                # 403, 404 y demas no hablan de la sesion.
                return True
            if intento == 1:
                desenlace = renovar_token(estado, cliente, persistir)
                if desenlace == "renovado":
                    continue
                if desenlace == "sin_respuesta":
                    # El token esta vencido pero no se pudo renovar por falta
                    # de red. Cerrar aca le haria perder la sesion por un corte
                    # momentaneo; se deja adentro y se reintenta en el proximo
                    # rerun, que es cuando probablemente ya haya internet.
                    return True
            return False
    return False


def licencia_vigente(
    estado,
    cliente,
    fingerprint: str,
    persistir: Callable[[], None] = _sin_efecto,
    ahora: Optional[float] = None,
    revalidar_horas: float = REVALIDAR_HORAS_POR_DEFECTO,
    gracia_dias: float = GRACIA_DIAS_POR_DEFECTO,
) -> bool:
    """Confirma la licencia, con cache y con tolerancia a la red.

    Tres reglas, en este orden:

    1. Si la ultima validacion sigue vigente, ni se sale a la red. Antes se
       validaba en CADA rerun de Streamlit -- decenas de llamadas por minuto --
       y cualquiera que fallara mandaba a la pantalla de activacion.
    2. Si el servidor rechaza, es definitivo: hay que activar.
    3. Si el servidor no contesta, vale la ultima validacion buena mientras
       este dentro de la gracia. Una licencia activada no deja de estarlo
       porque hoy no haya internet.
    """
    ahora = time.time() if ahora is None else ahora
    vigente_hasta = float(estado.get("license_valid_until") or 0)
    if estado.get("license_validated") and ahora < vigente_hasta:
        return True

    try:
        cliente.validate_license(estado.get("auth_token"), fingerprint)
    except ServidorInalcanzable:
        ultima = float(estado.get("license_last_check") or 0)
        if estado.get("license_validated") and ultima:
            return (ahora - ultima) < gracia_dias * 86400
        # Nunca se valido en este equipo: no hay nada que honrar.
        return False
    except RespuestaDelServidor as err:
        if err.es_sesion_invalida:
            # 401 habla del TOKEN, no de la licencia. Mandar a activar aca le
            # pediria el codigo a alguien cuya licencia esta perfecta, solo
            # porque se le vencio el token. De la sesion se ocupa la otra capa.
            ultima = float(estado.get("license_last_check") or 0)
            return bool(estado.get("license_validated") and ultima)
        estado["license_validated"] = False
        estado.pop("license_last_check", None)
        estado.pop("license_valid_until", None)
        persistir()
        return False

    estado["license_validated"] = True
    estado["license_last_check"] = ahora
    estado["license_valid_until"] = ahora + revalidar_horas * 3600
    persistir()
    return True


def marcar_licencia_activada(
    estado,
    persistir: Callable[[], None] = _sin_efecto,
    ahora: Optional[float] = None,
    revalidar_horas: float = REVALIDAR_HORAS_POR_DEFECTO,
) -> None:
    """Deja anotado que la licencia quedo activa recien."""
    ahora = time.time() if ahora is None else ahora
    estado["license_validated"] = True
    estado["license_last_check"] = ahora
    estado["license_valid_until"] = ahora + revalidar_horas * 3600
    persistir()
