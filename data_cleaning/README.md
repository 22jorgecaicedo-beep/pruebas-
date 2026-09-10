# Tarea de limpieza de datos (Lab Space Discussion — cap. 2.4)

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `limpieza_datos.ipynb` | El notebook. Súbelo a Google Colab y ejecútalo de arriba a abajo. |
| `limpieza_datos.py` | El mismo código como script, por si lo quieres leer o correr fuera de Colab. |
| `discussion_post.md` | El texto para pegar en el foro de D2L. |
| `peer_replies.md` | Las dos respuestas a compañeros que pide la actividad. |

## Cómo usarlo

1. Descarga tu dataset desde D2L (yo dejé el código apuntando a `spotify-2023.csv`).
2. Entra a [colab.research.google.com](https://colab.research.google.com) → **File → Upload notebook** → sube `limpieza_datos.ipynb`.
3. En el panel izquierdo, ícono de carpeta 📁 → botón de subir → sube tu `.csv`.
4. Si usaste otro dataset, cambia esta línea de la primera celda:
   ```python
   FILENAME = "spotify-2023.csv"
   ```
5. **Runtime → Run all**.
6. Toma captura de pantalla de la salida de los pasos 2, 4 y 9 (ahí se ve el antes y después).
7. Abre `discussion_post.md`, reemplaza cada `[X]` con los números reales que te imprimió, y pégalo en D2L.

## Nota

El código funciona con los cuatro datasets de la lista (`iris.csv`,
`movieprofit.csv`, `spotify-2023.csv`, `students.csv`): cada paso comprueba si la
columna existe antes de tocarla. Probado contra `iris.csv` y contra el formato de
`spotify-2023.csv`. Los pasos específicos de Spotify (la fecha de lanzamiento,
las comas en los números) simplemente se saltan si el dataset no tiene esas
columnas.
