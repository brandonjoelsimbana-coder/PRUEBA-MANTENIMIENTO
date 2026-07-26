# Asistente Excel IA — ESPE Latacunga

Aplicación web para Vercel que utiliza un modelo de OpenAI para interpretar preguntas en español y un motor determinista para calcular respuestas sobre cuatro archivos Excel.

## Publicación en Vercel

1. Descomprime este proyecto.
2. Sube la carpeta a un repositorio de GitHub.
3. En Vercel selecciona **Add New → Project** e importa el repositorio.
4. En **Environment Variables** agrega:
   - `OPENAI_API_KEY`: tu clave completa de OpenAI (marcada como secreta).
   - `OPENAI_MODEL`: `gpt-5.6-sol`.
5. Pulsa **Deploy**.

No coloques la clave en archivos, en GitHub ni en el código.

## Ejecución local opcional

```bash
npm install
cp .env.example .env.local
npm run dev
```

Edita `.env.local` y coloca allí la clave. Este archivo está excluido de Git.

## Cómo funciona

- El modelo interpreta la intención, columnas, filtros y operación solicitada.
- El servidor ejecuta conteos, promedios, sumas, listas y agrupaciones.
- Los cuatro Excel se consultan automáticamente.
- Si las fuentes presentan cantidades diferentes, se informa la discrepancia.
- La clave de OpenAI solo existe en el servidor.
- `next.config.mjs` obliga a Vercel a incluir los cuatro Excel dentro de la
  función `/api/ask`.
