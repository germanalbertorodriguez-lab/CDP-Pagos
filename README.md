# Pagos CDP — Instrucciones de configuración

## Pasos para activar la app

### 1. Configurar el Google Sheet

Abrí el sheet **Pagos CDP** y verificá que existan estas pestañas:
- `Equipos` — con columnas: EquipoID, NombreEquipo, Categoria, ...
- `CargaComprobantes` — donde se guardan los comprobantes
- `Jugadores` — con columnas: JugadorID, EquipoID, NombreJugador, MAS35, Intercategoria
- `Parametros` — fila 2: GLOBAL | PIN_ADMIN (el número 1603 que ya tenés)

### 2. Publicar el Apps Script

1. Abrí el sheet → **Extensiones → Apps Script**
2. Pegá el contenido de `AppScript.js` (reemplazá todo)
3. Guardá (Ctrl+S)
4. Ejecutá **`setupPINs`** una vez (botón ▶ con esa función seleccionada)
   - Esto agrega la columna PIN en la pestaña Equipos
   - Todos los equipos quedan con PIN `1234` por defecto
   - Podés cambiar cada uno en la columna I del sheet
5. Ejecutá **`setupFileIdColumn`** una vez
6. Clic en **Implementar → Nueva implementación**
   - Tipo: **Aplicación web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
7. Copiá la URL que aparece (algo como `https://script.google.com/macros/s/AKfy.../exec`)

### 3. Configurar el index.html

Abrí `index.html` y en la sección CONFIG (línea ~530) pegá la URL:

```javascript
const CONFIG = {
  SCRIPT_URL: 'https://script.google.com/macros/s/TU_URL_AQUI/exec',
  ADMIN_PIN: '1603',
  ...
};
```

### 4. Subir a Netlify

Subí el `index.html` a Netlify (igual que la app CDP).

---

## Cómo funciona

### Delegado
- Selecciona su equipo → ingresa su PIN (columna I del sheet Equipos)
- Ve la lista de jugadores con el estado de cada comprobante del mes
- Puede cargar un comprobante por jugador por mes (foto o PDF, máx 10MB)
- Si es rechazado, puede volver a cargar el comprobante de ese mes

### Administrador
- Ingresa con el PIN de la fila 2 de la pestaña Parametros (actualmente: 1603)
- Ve un resumen por equipo y mes
- Puede aprobar o rechazar cada comprobante con observación
- Los archivos se guardan automáticamente en Google Drive en la carpeta `PagosComprobantes_CDP/`

### Estructura Drive
```
PagosComprobantes_CDP/
  Abogados_A/
    Febrero/
      abc123_42_Febrero.jpg
      def456_17_Febrero.pdf
  Contadores_Z/
    Febrero/
      ...
```

---

## Cambiar PINs de delegados

En el Google Sheet, pestaña **Equipos**, columna **I (PIN)**. 
Cambiá el PIN de cada equipo según necesites.
