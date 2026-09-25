# Etiquetas Patagonia

App para imprimir etiquetas de 15 x 30 mm a partir de un código de barras, cruzando la
planilla de equivalencia (código → SKU, color, talle) con la planilla de STOCK (precio
RETAIL sin IVA). En la etiqueta figuran el precio sin IVA y el mismo con IVA 21%.

## Uso

1. Subí la planilla de equivalencia (`.xls` o `.xlsx`). Queda guardada hasta que cargues otra.
2. Subí la planilla de STOCK (`.xlsx`, solapa STOCK, columna RETAIL = precio sin IVA).
3. Escaneá o pegá el código de barras. La etiqueta se arma sola.

## Deploy en Portainer (Docker Compose)

1. Subí este repo a GitHub.
2. En Portainer: **Stacks → Add stack → Repository**, apuntá al repo de GitHub y a
   `docker-compose.yml`.
3. Deploy. La app queda escuchando en el puerto `8020` del host (configurable en el
   `docker-compose.yml`).
4. Los datos (la planilla grande cacheada) se guardan en el volumen `equivalencia-data`,
   así que sobreviven a reinicios y updates del contenedor.

## Deploy manual con Docker

```bash
docker build -t equivalencia-app .
docker run -d --name equivalencia-app -p 8020:8000 -v equivalencia-data:/app/data equivalencia-app
```

## Desarrollo local

```bash
pip install -r requirements.txt
python app.py
```

La app queda en `http://localhost:8000`.
