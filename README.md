# API Bifarma - Fork stateless multi-tenant

Fork de la API Bifarma de Oscar Garcia adaptado para servir a multiples farmacias
desde una sola instancia desplegada.

## Diferencia con el original

El codigo original usa un singleton global `_client = BifarmaClient()` que lee
`BIFARMA_USER` y `BIFARMA_PASSWORD` del entorno del servidor. Eso limita el servicio
a una sola farmacia por instancia.

Este fork convierte `get_client` en una factoria FastAPI: cada request crea un
`BifarmaClient` fresco con las credenciales del portal Bifarma que el cliente envie
en headers HTTP. El servidor no guarda sesion ni credenciales entre requests.

## Variables de entorno del servidor

Solo se necesita una variable:

| Variable | Descripcion |
|----------|-------------|
| `API_KEY` | Clave que deben incluir todos los clientes en el header `x-api-key`. |

Copia `.env.example` a `.env` y rellena el valor antes de arrancar en local.

`PORT` lo inyecta Render automaticamente.

## Headers por request

Cada llamada a cualquier endpoint (excepto `/status`) necesita tres headers:

| Header | Descripcion |
|--------|-------------|
| `x-api-key` | Clave de servicio (configurada en el servidor). |
| `x-bifarma-user` | Usuario del portal Bifarma de la farmacia. |
| `x-bifarma-password` | Contrasena del portal Bifarma de la farmacia. |

Si faltan `x-bifarma-user` o `x-bifarma-password` el endpoint devuelve `401`.

## Despliegue en Render

El fichero `render.yaml` configura el servicio (plan free, Python, puerto dinamico).
Para que Playwright funcione correctamente hay que usar `render-build.sh` como
`buildCommand` en lugar del `pip install` por defecto:

1. En el panel de Render, edita el servicio y cambia `Build Command` a:
   ```
   bash render-build.sh
   ```
2. Anade la variable de entorno `API_KEY` en la seccion "Environment".
3. Despliega.

El endpoint `/status` devuelve `"stateless": true` cuando el fork esta activo.

## Correr en local

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt -r requirements-dev.txt
cp .env.example .env   # editar API_KEY
uvicorn main:app --reload
```

## Tests

```bash
# desde la raiz de este repo, con el venv activo
pytest tests/test_stateless.py -v
```

Los 6 tests verifican el contrato stateless:
- no existe el singleton `_client`
- `get_client` exige ambos headers o lanza 401
- cada llamada crea una instancia independiente
- `/status` no expone `bifarma_user_configured` ni `bifarma_password_configured`
- las credenciales llegan al cliente sin modificacion
