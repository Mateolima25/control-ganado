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
- Registrar traslados entre lotes sin agregar un pesaje falso.
- Registro de campo para celular: muertes, enfermedades/observaciones, tratamientos y traslados, con fecha, lote confirmado, notas y foto opcional.
- Buscar animales por caravana y adjuntar una foto desde la cámara del teléfono; las fotos se reducen para guardarse con los datos sincronizados.
- Reconocer automáticamente los números de una caravana fotografiada, con confirmación manual antes de guardar el evento.
- Cambiar caravana conservando el identificador histórico de cada pesaje.
- Ganancia total y ganancia diaria.
- Importar Excel/CSV de alimentación.
- Detectar Fecha, Comederos, Animales, Total de comida, Sal, Bolsas y Días transcurridos.
- Calcular kg de comida por animal por día.
- Dashboard y gráficos.
- Datos sincronizados con Supabase y cache local para la cuenta propietaria.

## Identificación desde el celular

La caravana de la foto de referencia muestra un número impreso y no un QR/código de barras visible. En **Campo**, el OCR intenta leer los números de la foto en el propio dispositivo y buscar una caravana activa coincidente. Antes de guardar, se debe confirmar que el animal y el lote sean correctos; la lectura puede fallar si la foto está borrosa, lejos o mal iluminada. También se puede elegir un número detectado o buscarlo manualmente. El primer escaneo necesita conexión a internet para cargar el motor y los datos de reconocimiento; las fotos se procesan en el navegador y no se envían a ese servicio.

El botón negro de la caravana podría contener un chip RFID, pero debe confirmarse con el proveedor. RFID identifica un chip mediante radiofrecuencia y requiere un lector compatible; un teléfono común no puede leer por sí solo todos los chips RFID ganaderos. La lectura RFID se puede integrar más adelante si se conoce el estándar y el lector disponible.

## Formato de alimentación

La aplicación reconoce, entre otros, estos encabezados:

Fecha | Comederos | Animales (cantidad) | Total de comida | Sal | Bolsas de comida | DÍAS TRANSCURRIDOS

## Compartir lotes con acceso de solo lectura

La aplicación usa Supabase Auth y una fila JSON compartida con políticas RLS. Los usuarios con rol `viewer` pueden leer los lotes; solo el usuario `owner` puede modificarlos.

1. Crea un proyecto en Supabase y ejecuta el contenido de [`supabase-setup.sql`](supabase-setup.sql) en SQL Editor.
2. En Authentication, invita tu correo y los correos del grupo. La aplicación no ofrece registro público; solo podrán entrar las cuentas incluidas en `herd_members`.
3. Copia el UUID de cada usuario desde Authentication > Users y asígnale un rol en SQL Editor:

	```sql
	insert into public.herd_members (user_id, role)
	values ('UUID-DEL-PROPIETARIO', 'owner');

	insert into public.herd_members (user_id, role)
	values ('UUID-DEL-LECTOR', 'viewer');
	```

4. En Project Settings > API, copia la Project URL y la clave pública `anon` (nunca uses `service_role` en el frontend).
5. Configura `VITE_SUPABASE_URL` y `VITE_SUPABASE_ANON_KEY` como variables de entorno tanto en Vercel como en `.env.local` para desarrollo. Puedes tomar los nombres de [`.env.example`](.env.example). Después de agregar variables en Vercel, vuelve a desplegar.
6. Inicia sesión como propietario desde el navegador que actualmente tiene los lotes. Si la tabla compartida todavía no tiene datos, la aplicación copiará automáticamente allí los datos locales de ese navegador.
7. Cada miembro entra con su cuenta invitada. Los cambios del propietario aparecen para los lectores al actualizarse la vista, como máximo en unos 15 segundos.

Configura también la URL de tu sitio en Authentication > URL Configuration de Supabase y permite la URL de Vercel y `http://localhost:5173` como redirect URLs.

## Próxima etapa

Conectar el archivo de Google Drive/Sheets para leer la alimentación directamente desde la nube.
