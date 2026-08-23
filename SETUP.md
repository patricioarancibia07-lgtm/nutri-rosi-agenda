# Nutri-Rosi — Agenda online + recordatorios automáticos

Sistema de reserva de horas con confirmación y recordatorio automático por email, corriendo 100% gratis sobre Google Sheets + Google Apps Script. No necesita servidor ni hosting propio.

## Qué incluye

- **Code.gs** — toda la lógica: horarios disponibles, guardar reservas, enviar emails.
- **Index.html** — la página de reserva que verán tus pacientes.
- Los datos se guardan directo en una Hoja de cálculo de Google (pestaña "Reservas"), y la configuración del centro (servicios, horario, etc.) se edita en la pestaña "Config".

## Instalación (10–15 minutos, una sola vez)

1. **Crea la hoja de cálculo.** Ve a [sheets.google.com](https://sheets.google.com) → hoja en blanco. Ponle de nombre, por ejemplo, "Nutri-Rosi — Agenda".

2. **Abre el editor de scripts.** Menú *Extensiones → Apps Script*.

3. **Pega el código.**
   - En el archivo `Code.gs` que se abre por defecto, borra todo el contenido y pega el contenido del archivo `Code.gs` que te entregué.
   - Crea un archivo nuevo: ícono `+` junto a "Archivos" → *HTML*. Nómbralo exactamente `Index` (sin extensión). Borra el contenido de ejemplo y pega el de `Index.html`.
   - Guarda (Ctrl+S / Cmd+S).

4. **Ejecuta la configuración inicial.**
   - Arriba, donde dice la función a ejecutar, selecciona `configurarHojas` y presiona ▶ Ejecutar.
   - La primera vez te pedirá autorización: elige tu cuenta → "Avanzado" → "Ir a [nombre del proyecto] (no seguro)" → Permitir. Esto es normal, es tu propio script.
   - Esto crea automáticamente las pestañas **Reservas** y **Config** en tu hoja de cálculo.

5. **Completa la pestaña "Config"** en la hoja de cálculo (no en el editor de código) con los datos reales de tu centro: nombre, email de contacto, servicios que ofreces, duración de cada cita, días y horario de atención, y con cuántas horas de anticipación quieres que se envíe el recordatorio.

6. **Publica la página web.**
   - En el editor de Apps Script: *Implementar → Nueva implementación*.
   - Tipo: *Aplicación web*.
   - "Ejecutar como": Yo (tu cuenta).
   - "Quién tiene acceso": Cualquier usuario.
   - Presiona *Implementar* y copia la URL que te entrega. Esa es la página de reservas.

7. **Activa los recordatorios automáticos.**
   - Vuelve a seleccionar la función `crearTriggerRecordatorios` en el menú de funciones y presiona ▶ Ejecutar una sola vez.
   - Desde ese momento, cada hora el sistema revisa las reservas y envía el recordatorio a quienes tengan su cita dentro de las próximas X horas (lo que definas en Config), sin que tengas que hacer nada más.

## Cómo subirla a tu web

Tienes dos opciones simples:

- **Botón o link directo:** agrega un botón "Reservar hora" en tu sitio que apunte a la URL que copiaste en el paso 6.
- **Incrustada dentro de tu página:** pega esto en el HTML de tu sitio, reemplazando la URL:

  ```html
  <iframe src="TU_URL_AQUI" width="100%" height="720" style="border:none;"></iframe>
  ```

## Cómo se ven las reservas

Todas las reservas quedan registradas en la pestaña "Reservas" de tu hoja de cálculo, con nombre, email, teléfono, servicio, fecha, hora y estado. Si necesitas cancelar una, simplemente escribe "Cancelada" en la columna "Estado" de esa fila — el horario queda libre de nuevo y no se le enviará recordatorio.

## Si algo falla

- **"Faltan datos obligatorios" o el botón de reservar no se activa:** revisa que hayas elegido un horario (los horarios aparecen como botones después de elegir la fecha).
- **No llegan los correos:** confirma que ejecutaste `crearTriggerRecordatorios` (para los recordatorios) y que el email de confirmación se manda solo al crear la reserva — revisa la carpeta de spam del destinatario.
- **La página se ve en blanco:** vuelve a *Implementar → Administrar implementaciones* y confirma que el acceso está en "Cualquier usuario".

## Dashboard de profesionales (nuevo)

Cada profesional (nutricionista, kinesiólogo) puede entrar a un dashboard propio para ver sus horas, marcarlas como atendidas, cancelarlas, dejar notas, y bloquear horarios en los que no atiende (vacaciones, hora personal, etc).

### Instalar el dashboard (una sola vez, sin borrar tus reservas)

1. En el editor de Apps Script, reemplaza `Código.gs` por la versión más reciente que te entregué, y agrega el archivo nuevo `Dashboard.html` (ícono `+` → HTML → nómbralo exactamente `Dashboard`, pega el contenido). Guarda.
2. Selecciona la función `agregarColumnaSucursal` y ▶ Ejecuta (si ya la corriste antes, no pasa nada por repetirla).
3. Selecciona la función `configurarDashboard` y ▶ Ejecuta. Esto crea la pestaña **Servicios** (dice qué profesional atiende cada servicio — revísala y ajústala si hace falta) y la pestaña **Usuarios** (vacía por ahora), y agrega las columnas "Profesional" y "Nota" a Reservas.
4. Crea una cuenta por cada profesional:
   - Abre la función `crearUsuario` en el código.
   - Edita las 4 líneas de arriba (usuario, contraseña, rol, nombre). El `rol` debe ser exactamente `nutricionista`, `kinesiologo`, o `admin` (admin ve y gestiona las horas de todos).
   - ▶ Ejecuta. Repite este paso (editar y ejecutar) una vez por cada persona.
5. Vuelve a implementar el proyecto como nueva versión (*Implementar → Administrar implementaciones → editar → Nueva versión → Implementar*) para que la web tome todos estos cambios.

### Cómo entran los profesionales

La URL del dashboard es la misma de siempre, agregando `?page=dashboard` al final. Por ejemplo:

```
https://script.google.com/macros/s/TU_ID/exec?page=dashboard
```

Cada profesional entra ahí con el usuario/contraseña que le creaste. Ve solo sus propias horas (según el Rol que le pusiste); si creaste una cuenta `admin`, esa ve y gestiona las horas de todos.

**Nota sobre seguridad:** este es un sistema simple, adecuado para un equipo chico — las contraseñas se guardan encriptadas (nunca en texto plano), pero no tiene la robustez de un sistema con cuentas de Google reales. No lo uses para información más sensible que la agenda.

## Próximos módulos

Esto cubre **Agenda online + recordatorios automáticos** y el **dashboard de profesionales**. Cuando quieras seguimos con el resto de la lista (WhatsApp, Mercado Pago, ficha clínica, CRM, estadísticas) — cada uno como su propio módulo, para no mezclar todo en un solo sistema difícil de mantener.
