# Asistente Excel — ESPE Latacunga

Versión estable para Vercel sin OpenAI ni servicios de pago. Conserva el diseño,
el título institucional y la carga conjunta de los cuatro Excel. El motor local
interpreta preguntas con distintas palabras, singular/plural, filtros, conteos,
listas, sumas, promedios, máximos, mínimos, comparaciones y agrupaciones.

## Publicación en Vercel

1. Descomprime este proyecto.
2. Sube la carpeta a un repositorio de GitHub.
3. En Vercel selecciona **Add New → Project** e importa el repositorio.
4. Pulsa **Deploy**. No se requieren variables de entorno.

## Ejecución local opcional

```bash
npm install
npm run dev
```

## Cómo funciona

- Los cuatro Excel se cargan desde `public/data` directamente en la página.
- No utiliza OpenAI, claves, saldo ni variables de entorno.
- Consulta automáticamente todas las fuentes relevantes, sin pedir al usuario
  que seleccione un libro.
- Incluye reglas corregidas para técnicos, asesores, lavado, control de calidad,
  OT, vehículos, estados, costos, horas, capacidad, cumplimiento y agrupaciones.
- Detecta diferencias entre archivos y las muestra en vez de elegir un valor al
  azar.
- Cada respuesta con datos numéricos ofrece gráficas de barras, línea y circular.
- La pregunta **¿Qué puedo preguntar?** muestra un catálogo de ejemplos basado
  en los campos disponibles.
