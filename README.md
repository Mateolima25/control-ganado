# Control de Ganado

Aplicación React/Vite para controlar lotes, pesajes de True-Test y alimentación.

## Instalar

Abrir esta carpeta en VS Code y ejecutar:

```bash
npm install
npm run dev
```

Luego abrir la dirección que muestre Vite, normalmente:

http://localhost:5173

## Qué incluye

- Crear lotes.
- Importar Excel/CSV de True-Test.
- Detectar columnas EID, RFID, VID, Caravana, Peso y Fecha.
- Comparar animales por número de caravana.
- Historial de pesajes.
- Ganancia total y ganancia diaria.
- Importar Excel/CSV de alimentación.
- Detectar Fecha, Comederos, Animales, Total de comida, Sal, Bolsas y Días transcurridos.
- Calcular kg de comida por animal por día.
- Dashboard y gráficos.
- Datos guardados en el navegador mediante localStorage.

## Formato de alimentación

La aplicación reconoce, entre otros, estos encabezados:

Fecha | Comederos | Animales (cantidad) | Total de comida | Sal | Bolsas de comida | DÍAS TRANSCURRIDOS

## Próxima etapa

Conectar el archivo de Google Drive/Sheets para leer la alimentación directamente desde la nube y agregar usuarios/base de datos para que la información no dependa del navegador.
