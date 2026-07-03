import pytest
from fastapi import HTTPException
from fastapi.testclient import TestClient

import main

API_KEY = "test-api-key"


@pytest.fixture(autouse=True)
def _set_api_key(monkeypatch):
    # require_api_key lee API_KEY del entorno; sin esto responde 500.
    monkeypatch.setenv("API_KEY", API_KEY)


@pytest.fixture
def client():
    return TestClient(main.app)


def test_health_no_requiere_api_key(client):
    # /health debe responder 200 sin ninguna cabecera de autenticacion.
    r = client.get("/health")
    assert r.status_code == 200
    body = r.json()
    assert body["ok"] is True
    assert "version" in body


def test_no_hay_cliente_global():
    # El singleton compartido entre farmacias debe desaparecer.
    assert not hasattr(main, "_client")


def test_get_client_exige_headers_bifarma():
    with pytest.raises(HTTPException) as exc:
        main.get_client(None, None)
    assert exc.value.status_code == 401


def test_get_client_crea_instancia_por_request():
    a = main.get_client("u1", "p1")
    b = main.get_client("u2", "p2")
    assert a is not b                 # instancia fresca por request
    assert a.session is not b.session  # sin sesion compartida entre farmacias
    assert a.username == "u1" and a.password == "p1"


def test_status_no_expone_credenciales_bifarma(client):
    r = client.get("/status", headers={"x-api-key": API_KEY})
    assert r.status_code == 200
    body = r.json()
    assert "bifarma_user_configured" not in body
    assert "bifarma_password_configured" not in body


def test_login_401_sin_headers_bifarma(client):
    # api-key correcto pero sin credenciales Bifarma -> lo corta get_client.
    r = client.post("/login", headers={"x-api-key": API_KEY})
    assert r.status_code == 401
    assert "bifarma" in r.json()["detail"].lower()


def test_credenciales_llegan_al_cliente(client, monkeypatch):
    recibido = {}

    def fake_login(self, force=False):
        recibido["username"] = self.username
        recibido["password"] = self.password
        return {"logged_in": True, "reused_session": False}

    monkeypatch.setattr(main.BifarmaClient, "ensure_login", fake_login)
    r = client.post(
        "/login",
        headers={
            "x-api-key": API_KEY,
            "x-bifarma-user": "apicornell",
            "x-bifarma-password": "secreto",
        },
    )
    assert r.status_code == 200
    assert recibido == {"username": "apicornell", "password": "secreto"}
