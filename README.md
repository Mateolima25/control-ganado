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
- Cambiar caravana conservando el identificador histórico de cada pesaje.
- Ganancia total y ganancia diaria.
- Importar Excel/CSV de alimentación.
- Detectar Fecha, Comederos, Animales, Total de comida, Sal, Bolsas y Días transcurridos.
- Calcular kg de comida por animal por día.
- Dashboard y gráficos.
- Datos sincronizados con Supabase y cache local para la cuenta propietaria.

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
