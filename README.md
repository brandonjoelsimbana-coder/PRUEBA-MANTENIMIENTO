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
  azar. Para indicadores de vehículos y tipos de trabajo utiliza
  `Entregable_3_Dashboard_Indicadores` como fuente principal y presenta los
  otros libros como control cruzado.
- Cada respuesta con datos numéricos ofrece gráficas de barras, línea y circular.
- Incluye un explorador con **474 preguntas y variantes** generadas desde los
  estados, técnicos, asesores, bahías, tipos de vehículo, mantenimientos y
  definiciones reales de los cuatro Excel. Todas son clicables y se pueden
  filtrar por texto.
- La pregunta **¿Qué puedo preguntar?** resume el catálogo por tema.
- Interpreta “cuántas horas trabajó cada técnico” como suma de horas, no como
  conteo de registros. Si se pregunta por el mes, aclara que los archivos
  disponibles cubren del 13 al 27 de julio de 2026.
- Reconoce variantes como camioneta/camionetas/pickup, SUV, automóvil/sedán,
  monovolumen/minivan y kilometrajes escritos como `20.000`, `20000` o
  `20 mil km`.

## Verificación

```bash
npm test
npm run build
```

Las pruebas validan los cuatro archivos, las 17 hojas, las consultas críticas
de personal, vehículos, órdenes y horas, además del tamaño del catálogo.
