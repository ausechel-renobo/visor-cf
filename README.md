# Visor CF · Tablero de Gerencia RENOBO 2026–2027

Visor de los 15 KPIs de la agenda gerencial. **Consulta y edición** en una página
estática; los datos viven en una hoja de Google Sheets del gerente.

- Sin servidor, sin Apps Script, sin costo.
- La hoja queda **privada**: el acceso va autenticado con la cuenta de Google del gerente.
- Si el visor fallara alguna vez, la hoja sigue ahí, legible y editable. Ningún dato
  queda atrapado en un formato propietario.

```
GitHub Pages  ──API de Sheets (OAuth, desde el navegador)──▶  Google Sheet
  el visor                                                    la base de datos
```

---

## Puesta en marcha

### Paso 0 · Credenciales de Google (una vez, ~20 min)

Es el único punto donde algo puede no funcionar: **algunos tenants corporativos de
Google Workspace restringen las aplicaciones OAuth de terceros.** Conviene hacerlo
antes que nada.

1. Entra a [console.cloud.google.com](https://console.cloud.google.com) y crea un
   proyecto. No requiere facturación.
2. **APIs y servicios → Biblioteca** → busca *Google Sheets API* → **Habilitar**.
3. **APIs y servicios → Pantalla de consentimiento de OAuth**:
   - Si el proyecto está dentro del Workspace de RENOBO, elige **Interno**. Es el
     escenario ideal: sin verificación de Google y sin advertencias.
   - Si es una cuenta personal, elige **Externo**, déjalo en modo *prueba* y añade
     el correo del gerente en **Usuarios de prueba**. Funciona igual, pero la
     primera vez aparece *«Google no ha verificado esta aplicación»* y hay que
     entrar por **Avanzado → Continuar**.
   - Alcance requerido: `https://www.googleapis.com/auth/spreadsheets`.
4. **Credenciales → Crear credenciales → ID de cliente de OAuth**, tipo
   **Aplicación web**. En *Orígenes de JavaScript autorizados* añade:
   - `https://ausechel-renobo.github.io`
   - `http://localhost:8080` (opcional, para desarrollo)

   > Es el **origen**, no la URL del visor: sin `/visor-cf` y sin barra final.
   > Google rechaza el valor si lleva ruta. El origen es el mismo para todos los
   > repositorios de la cuenta, así que basta con registrarlo una vez.

   Copia el **ID de cliente** — termina en `.apps.googleusercontent.com`.

> El ID de cliente **no es un secreto**: es público por diseño y solo funciona
> desde los orígenes autorizados. Sin la sesión de Google del gerente no da acceso
> a nada, así que puede vivir en un repositorio público.

### Paso 1 · Publicar el visor

```bash
git init && git add . && git commit -m "Visor CF"
git remote add origin https://github.com/ausechel-renobo/visor-cf.git
git push -u origin main
```

En **Settings → Pages** del repositorio, elige *Deploy from a branch* → `main` / `/ (root)`.
Queda publicado en `https://ausechel-renobo.github.io/visor-cf/`.

Para publicar una mejora más adelante basta con `git push`: el gerente solo tiene que
recargar. Los datos no se tocan, porque viven en la hoja y no en la app.

### Paso 2 · Primera conexión

El gerente abre la URL y verá la pantalla de configuración:

- **ID de cliente OAuth** — el del paso 0.
- **ID de la hoja de cálculo** — se puede pegar la URL completa de la hoja; el visor
  extrae el ID solo. **Si se deja vacío, el visor crea la hoja** con las cinco
  pestañas y los 15 KPIs ya cargados desde el Excel original.

Ambos valores quedan guardados en su navegador; no se vuelven a pedir.

Por último, **Instalar aplicación** en Edge o Chrome (el icono ⊕ en la barra de
direcciones): queda con icono propio y ventana sin barra de navegador.

---

## Uso diario

| Acción | Qué hace |
|---|---|
| Editar una celda amarilla | Acumula el cambio. La cabecera muestra *«N cambios sin guardar»*. |
| **Guardar** (`Ctrl+S`) | Escribe todo en la hoja de una sola vez. |
| **Actualizar** | Vuelve a leer la hoja sin recargar la página. |
| **Cerrar mes** | Congela el avance de los 15 KPIs en la columna del mes y adelanta el corte. Construye el histórico solo. |
| **Hoja** | Abre la hoja de cálculo en otra pestaña. |
| **Comité** + `Ctrl+P` | PDF apaisado de una página para el comité directivo. |

Si se cierra la ventana con cambios pendientes, el navegador avisa; y al volver a
abrir, el visor ofrece recuperarlos.

---

## Estructura de la hoja

Cinco pestañas, encabezados en la fila 1, una fila por registro. **Es
almacenamiento, no interfaz**: sin fórmulas, sin formato condicional, sin celdas
combinadas — todo cálculo vive en el visor.

| Hoja | Columnas |
|---|---|
| `KPIs` | `id · pilar · nombre · unidad · avance · meta2026 · meta2027 · metaFinal · responsable · avanceDesdeHitos` |
| `Serie` | `kpiId` + una columna por mes, de `2026-08` a `2028-01` |
| `Hitos` | `kpiId · nombre · hecho · fecha` |
| `Bitacora` | `kpiId · fecha · autor · texto` |
| `Config` | `clave · valor` — `titulo`, `mesCorte`, `umbralAtencion`, `revision` |

Unidades admitidas en `KPIs`: `conteo`, `porcentaje` (decimal: 0,15 = 15 %),
`moneda` (millones de COP) y `m2`.

Los KPIs con `avanceDesdeHitos = TRUE` (12 y 13) calculan su avance contando los
hitos marcados; en la tabla aparecen como *auto* y no se editan a mano.

`revision` es un contador que el visor incrementa en cada guardado. Si la hoja se
edita por fuera y el contador queda desalineado, el visor avisa antes de
sobrescribir en vez de pisar los cambios a ciegas.

---

## Cálculos

Nada de esto se guarda: se recalcula en cada carga, así que el semáforo nunca queda
desfasado respecto a su dato.

| Indicador | Fórmula |
|---|---|
| Meta del año | 2026 → `meta2026`; 2027+ → `meta2027`, o la meta final si aún no está definida |
| Progreso | `avance / metaFinal`, acotado a [0 ; 1,2] — replica `G6:G20` del Excel |
| Estado | `≥ meta del año` → **En meta**; `≥ 40 %` → **Atención**; si no → **Grave** |
| Ritmo requerido | `(meta del año − avance) / meses que faltan del año` |
| Proyección | extrapolación lineal de los últimos tres puntos de la serie |

El Excel promediaba el avance contra la *meta final* (12,8 %), un número que
subestima la gestión de 2026 porque mezcla el horizonte de los dos años. El visor
muestra **avance vs. meta del año** junto a **avance vs. meta final**.

---

## Páginas de apoyo (no requieren Google)

Ambas se abren con doble clic y funcionan sin conexión ni credenciales.

- **`pruebas.html`** — verifica que el semáforo y el progreso que calcula el visor
  coinciden con las columnas `G` y `H` de `Dashboard.xlsx`, incluido el promedio de
  `G21` (12,842643 %). Son 31 comprobaciones. Conviene volver a abrirla cada vez que
  se toque `calculos` o la semilla.
- **`vista-previa.html`** — el visor con datos de ejemplo y una serie sintética, para
  revisar el diseño sin montar OAuth. Acepta
  `?vista=resumen|tabla|detalle|comite` y `&kpi=KPI-07`. Muestra cualquier error de
  JavaScript en una banda roja en lugar de dejar la página en blanco.

---

## Recuperación

- **Deshacer un guardado**: en la hoja, *Archivo → Historial de versiones*.
- **Empezar de cero**: borrar el ID de la hoja en la configuración del visor y dejarlo
  vacío; se crea una hoja nueva con la semilla original.
- **El visor no abre**: los datos están íntegros en la hoja. Se puede consultar y
  editar directamente ahí mientras se resuelve.
- **Restablecer la configuración local**: en la consola del navegador,
  `localStorage.clear()` y recargar.

---

## Pendientes de confirmar con la gerencia

1. **KPI 06 — Recursos FCO comprometidos**: avance 138 contra meta 115.000 (0,06 %).
   Parece una discrepancia de unidades. Se cargó tal cual está en el Excel y quedó
   anotado en su bitácora.
2. **Unidad monetaria**: se asumen **millones de COP** en los KPI 04, 06 y 09.
3. **Metas 2027**: vacías hasta que se definan; mientras tanto el semáforo de 2027
   usa la meta final.
4. Correcciones aplicadas respecto al original: el encabezado decía «14 KPIs» pero
   hay 15; el KPI 10 tenía formato decimal siendo porcentaje; y el responsable del
   KPI 14 decía «Sugerencia corporativa» (se normalizó a *Subgerencia corporativa*).
