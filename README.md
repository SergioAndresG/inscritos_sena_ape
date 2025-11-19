# ✨ Automatización de Ingreso de Inscritos en APE – SENA

<p align="center">
  <strong>Automatización inteligente para la inscripción masiva de aprendices en la Agencia Pública de Empleo del SENA</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white" alt="Python" />
  <img src="https://img.shields.io/badge/Selenium-4.x-43B02A?logo=selenium&logoColor=white" alt="Selenium" />
  <img src="https://img.shields.io/badge/Excel-Compatible-217346?logo=microsoftexcel&logoColor=white" alt="Excel" />
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white" alt="Windows" />
</p>

---

## 📋 Tabla de Contenidos

- [¿Qué es esta herramienta?](#-qué-es-esta-herramienta)
- [¿Por qué usarla?](#-por-qué-usarla)
- [Descargar e Instalar](#-descargar-e-instalar)
- [Primer Uso](#-primer-uso)
- [Cómo Usar la Aplicación](#-cómo-usar-la-aplicación)
- [Estructura del Archivo Excel](#-estructura-del-archivo-excel)
- [Preguntas Frecuentes](#-preguntas-frecuentes)
- [Solución de Problemas](#-solución-de-problemas)
- [Para Desarrolladores](#-para-desarrolladores)

---

## 🎯 ¿Qué es esta herramienta?

Esta aplicación es un **sistema de automatización robótico (RPA)** diseñado específicamente para el personal del APE que necesita registrar aprendices en la plataforma de la Agencia Pública de Empleo (APE).

El programa automatiza todo el proceso de ingreso de datos, desde el inicio de sesión hasta el registro completo de cada aprendiz, eliminando la necesidad de hacerlo manualmente uno por uno.

### 🎬 Demo Visual

```
┌─────────────────────────────────────────────────────┐
│  📊 Carga tu Excel                                  │
│  🔐 Ingresa tus credenciales                        │
│  ▶️  Inicia el proceso                              │
│  ☕ Toma un café mientras la app trabaja            │
│  ✅ Revisa el reporte final                         │
└─────────────────────────────────────────────────────┘
```

---

## 💡 ¿Por qué usarla?

### El problema que resuelve

El registro manual de aprendices colocados en APE es un proceso que:
- ⏰ **Consume hasta 2 horas** cuando hay múltiples registros
- 🔄 **Es repetitivo y tedioso**, requiriendo los mismos pasos una y otra vez
- ⚠️ **Es propenso a errores** por digitación o copiar datos incorrectos
- 📊 **Dificulta el seguimiento** de qué registros ya fueron procesados

### La solución

Esta herramienta convierte un proceso de **2 horas en solo 25 minutos**, logrando una **mejora del 83%** en eficiencia:

<table>
<tr>
<td align="center" width="33%">

### ⚡ Velocidad
Procesa múltiples registros automáticamente sin intervención manual

</td>
<td align="center" width="33%">

### 🎯 Precisión
Elimina errores de digitación al tomar datos directamente del Excel

</td>
<td align="center" width="33%">

### 📝 Trazabilidad
Genera logs detallados de cada operación para auditoría

</td>
</tr>
</table>

---

## 📥 Descargar e Instalar

### Para Usuarios Finales (Recomendado)

> ✅ **No necesitas instalar Python, Git ni ningún programa adicional**  
> Todo está incluido en un único archivo ejecutable.

#### 1️⃣ Descargar la Aplicación

**Opción A: Descarga directa**
1. Ve a la sección [**📦 Releases**](https://github.com/SergioAndresG/inscritos_sena_ape/releases/latest)
2. Descarga el archivo más reciente: `SENA_Automation_App_vX.X.X.zip`
3. Extrae el contenido en una carpeta de tu preferencia

**Opción B: Descarga rápida**
```
📥 Descarga directa: [Última versión estable]
```

#### 2️⃣ Contenido del Paquete

Después de extraer el `.zip`, encontrarás:

```
📁 SENA_Automation_App/
├── 📄 SENA_Automation_App.exe          ← Archivo principal (ejecutar este)
├── 📁 funciones_formularios/           ← Módulos de la aplicación
├── 📁 funciones_loggs/                 ← Sistema de registro
├── 📁 perfilesOcupacionales/           ← Mapeo de programas
├── 📁 Iconos/                          ← Recursos gráficos
├── 📁 URLS/                            ← Configuración de enlaces
├── 📄 README.md                        ← Este archivo
└── 📄 plantilla_inscritos.xls         ← Plantilla Excel ejemplo
```

#### 3️⃣ Ubicación Recomendada

Te sugerimos colocar la aplicación en una carpeta dedicada:

```
📁 C:\Usuarios\TuNombre\Documentos\
   └── 📁 Automatizacion_APE_SENA\
       ├── 📄 SENA_Automation_App.exe
       ├── 📁 funciones_formularios\
       ├── 📁 funciones_loggs\
       ├── 📁 Logs\                     (se crea automáticamente)
       └── 📁 config\                   (se crea automáticamente)
```

#### 4️⃣ Primera Ejecución

1. **Doble clic** en `SENA_Automation_App.exe`
2. **Si Windows Defender muestra una advertencia:**
   
   ```
   ⚠️ "Windows protegió tu PC"
   ```
   
   - Haz clic en **"Más información"**
   - Luego clic en **"Ejecutar de todas formas"**

3. **¿Por qué aparece esta advertencia?**
   - El archivo no tiene firma digital (certificado que cuesta dinero)
   - Windows no reconoce aplicaciones nuevas
   - Es normal y **seguro** (el código fuente está disponible para revisión)

---

## 🔐 Primer Uso

### Configuración Inicial (Solo la primera vez)

Al ejecutar la aplicación por primera vez, verás un asistente de configuración:

#### Paso 1: Credenciales de APE

<table>
<tr>
<td width="50%">

**📝 Información requerida:**
- Usuario - Numeró de documento (Debe ser un usuario con rol de funcionario)
- Contraseña de acceso a APE

</td>
<td width="50%">

**🔒 Seguridad:**
- Las credenciales se guardan **localmente** en tu equipo
- Se almacenan en: `config/credentials.json`
- **Nunca** se envían a servidores externos

</td>
</tr>
</table>

#### Paso 2: Mapeo de Perfiles Ocupacionales

La aplicación necesita relacionar los **programas de formación** con los **perfiles ocupacionales** de APE.

**¿Qué sucede?**
- Si el programa ya está en la base de datos → Se usa automáticamente
- Si el programa **NO** está registrado → La app te preguntará el perfil correcto

**Ejemplo:**
```
Programa: "TECNÓLOGO EN ANÁLISIS Y DESARROLLO DE SOFTWARE"
         ↓
Perfil:   "Desarrollador de Software"
```

Este mapeo se guarda en `perfilesOcupacionales/mapeo_programas.json` para futuros usos.

---

## 📖 Cómo Usar la Aplicación

### Paso 1: Preparar el Archivo Excel

Tu archivo Excel **DEBE** tener exactamente estas columnas (respeta mayúsculas y tildes):

| Columna | Tipo de Dato | Ejemplo | Observaciones |
|---------|--------------|---------|---------------|
| **Tipo de Documento** | Texto | CC, TI, CE, PEP | Siglas sin puntos |
| **Número de Documento** | Número | 1234567890 | Sin puntos, comas ni espacios |
| **Nombres** | Texto | Juan Carlos | Nombres completos |
| **Apellidos** | Texto | Pérez González | Apellidos completos |
| **Programa de Formación** | Texto | TECNÓLOGO EN ANÁLISIS Y DESARROLLO DE SOFTWARE | Nombre completo y exacto |
| **Empresa** | Texto | ABC Tecnología S.A.S. | Razón social completa |
| **NIT Empresa** | Número | 900123456 | Sin dígito de verificación |
| **Fecha de Ingreso** | Fecha | 15/01/2024 | Formato: DD/MM/AAAA |

**📄 Descarga la plantilla:** [`plantilla_inscritos.xls`](./plantilla_inscritos.xls)

### Paso 2: Ejecutar la Aplicación

1. **Abrir la aplicación:**
   - Doble clic en `SENA_Automation_App.exe`
   - Espera a que cargue la interfaz gráfica

2. **Cargar el archivo Excel:**
   - Clic en botón **"📂 Seleccionar archivo Excel"**
   - Busca y selecciona tu archivo `.xls` o `.xlsx`
   - La aplicación validará la estructura automáticamente

3. **Verificar configuración:**
   - Revisa que tus credenciales sean correctas
   - Confirma que los programas estén mapeados

4. **Iniciar automatización:**
   - Clic en botón **"▶️ Iniciar Proceso"**
   - Se abrirá Google Chrome automáticamente
   - **No cierres el navegador ni interactúes con él**

### Paso 3: Monitorear el Proceso

Durante la ejecución verás:

```
┌────────────────────────────────────────────┐
│  [████████████░░░░░░░░░░] 60%            │
│                                            │
│  📊 Registros procesados: 12/20            │
│  ✅ Exitosos: 11                           │
│  ❌ Fallidos: 1                            │
│                                            │
│  🔄 Procesando: PÉREZ, Juan Carlos         │
│  ⏱️ Tiempo estimado: 5 minutos             │
└────────────────────────────────────────────┘
```

**Opciones disponibles:**
- ⏸️ **Pausar**: Detiene temporalmente el proceso
- ⏹️ **Cancelar**: Detiene completamente (puedes reanudar después)
- 📋 **Ver Logs**: Muestra detalles técnicos en tiempo real

### Paso 4: Revisar Resultados

Al finalizar, la aplicación mostrará:

✅ **Resumen de Ejecución:**
```
╔════════════════════════════════════════════╗
║        PROCESO COMPLETADO                  ║
╠════════════════════════════════════════════╣
║  Total procesados:     20                  ║
║  ✅ Exitosos:          18                  ║
║  ❌ Fallidos:          2                   ║
║  ⏱️ Tiempo total:       14 minutos         ║
╚════════════════════════════════════════════╝
```

📄 **Archivos generados:**
- `Logs/registro_YYYYMMDD_HHMMSS.log` → Log detallado
- `Logs/errores_YYYYMMDD.xlsx` → Registros fallidos (si aplica)

---

## 📊 Estructura del Archivo Excel

### Validaciones Automáticas

Antes de iniciar, la aplicación verifica:

| Validación | Descripción |
|------------|-------------|
| ✔️ **Columnas requeridas** | Todas las columnas obligatorias deben existir |
| ✔️ **Celdas vacías** | No puede haber campos obligatorios vacíos |
| ✔️ **Formato de fechas** | Debe ser DD/MM/AAAA (ej: 15/01/2024) |
| ✔️ **Números de documento** | Solo números, sin puntos ni espacios |
| ✔️ **Tipos de documento** | Solo valores válidos: CC, TI, CE, PEP, etc. |
| ✔️ **NIT válido** | Solo números, sin dígito de verificación |

### Errores Comunes y Soluciones

| ❌ Error | ✅ Solución |
|---------|-----------|
| "Columna 'X' no encontrada" | Verifica el nombre exacto (mayúsculas y tildes) |
| "Fecha inválida en fila 5" | Usa formato DD/MM/AAAA (ej: 15/01/2024) |
| "Documento inválido" | Elimina puntos, guiones y espacios |
| "Celda vacía en columna 'Nombres'" | Completa todos los campos obligatorios |
| "Programa no encontrado" | La app te pedirá el perfil ocupacional |

### Ejemplo de Registro Válido

```excel
| CC | 1234567890 | Juan Carlos | Pérez González | TECNÓLOGO EN ANÁLISIS Y DESARROLLO DE SOFTWARE | Tech Solutions S.A.S. | 900123456 | 15/01/2024 |
```

---

## ❓ Preguntas Frecuentes

<details>
<summary><strong>¿Necesito instalar Python u otros programas?</strong></summary>

**No.** El archivo `.exe` es completamente autónomo e incluye todo lo necesario:
- Python embebido
- Todas las librerías (Selenium, pandas, CustomTkinter, etc.)
- ChromeDriver actualizado
- Módulos personalizados

Solo necesitas tener **Google Chrome** instalado en tu equipo.
</details>

<details>
<summary><strong>¿Funciona con otros navegadores como Firefox o Edge?</strong></summary>

Actualmente está optimizado solo para **Google Chrome**. Versiones futuras podrían incluir otros navegadores.
</details>

<details>
<summary><strong>¿Qué pasa si se interrumpe el proceso?</strong></summary>

La aplicación guarda el progreso automáticamente. Puedes:
1. Reanudar desde el último registro procesado
2. Revisar el archivo de logs para ver hasta dónde llegó
3. Eliminar los registros ya procesados del Excel y continuar
</details>

<details>
<summary><strong>¿Cuántos registros puedo procesar a la vez?</strong></summary>

**No hay límite técnico**, pero recomendamos:
- ✅ **1-50 registros**: Óptimo para revisión rápida
- ⚠️ **50-100 registros**: Recomendable dividir en lotes
- ❌ **+100 registros**: Dividir en archivos más pequeños

Procesar en lotes facilita la identificación de errores.
</details>

<details>
<summary><strong>¿Mis credenciales están seguras?</strong></summary>

**Sí.** Las credenciales:
- Se almacenan **solo en tu equipo** (carpeta `config/`)
- Usan encriptación básica
- **Nunca** se envían a internet (excepto a APE para login)
- No se comparten con servidores de terceros

**Recomendación:** No compartas la carpeta `config/` con otras personas.
</details>

<details>
<summary><strong>¿Puedo modificar el mapeo de programas manualmente?</strong></summary>

**Sí.** Puedes editar el archivo:
```
perfilesOcupacionales/mapeo_programas.json
```

**Formato:**
```json
{
  "NOMBRE DEL PROGRAMA": "Perfil Ocupacional",
  "TECNÓLOGO EN ANÁLISIS Y DESARROLLO DE SOFTWARE": "Desarrollador de Software"
}
```

Guarda y reinicia la aplicación.
</details>

<details>
<summary><strong>¿Qué hago si un registro falla?</strong></summary>

1. **Durante la ejecución:** La app continúa con los siguientes
2. **Al finalizar:** Revisa el reporte de errores
3. **En los logs:** Encuentra detalles específicos del error
4. **Corrección:** Ajusta los datos y vuelve a procesar solo ese registro

La aplicación genera un archivo Excel con los registros fallidos para facilitar su corrección.
</details>

<details>
<summary><strong>¿Puedo usar la app en varios equipos?</strong></summary>

**Sí.** Simplemente:
1. Copia la carpeta completa a otro equipo
2. Ejecuta el `.exe`
3. Configura las credenciales (son por equipo)

No hay límite de instalaciones.
</details>

---

## 🔧 Solución de Problemas

### Problema 1: Windows Defender bloquea el archivo

**Síntomas:**
- "Windows protegió tu PC"
- El archivo desaparece después de descargarlo
- Antivirus elimina el ejecutable

**Solución:**

**Paso A: Permitir ejecución única**
1. Clic en **"Más información"**
2. Clic en **"Ejecutar de todas formas"**

**Paso B: Agregar excepción permanente**
1. Abre **Windows Security** (Seguridad de Windows)
2. Ve a **"Protección contra virus y amenazas"**
3. Clic en **"Administrar configuración"**
4. Desplázate hasta **"Exclusiones"**
5. Clic en **"Agregar o quitar exclusiones"**
6. **"Agregar una exclusión"** → **"Carpeta"**
7. Selecciona la carpeta donde está el `.exe`

**Paso C: Desbloquear archivo descargado**
1. Clic derecho en `SENA_Automation_App.exe`
2. **Propiedades**
3. En la pestaña **General**, marca ☑️ **"Desbloquear"**
4. **Aplicar** → **Aceptar**

---

### Problema 2: Error de ChromeDriver o navegador

**Síntomas:**
```
❌ Error: ChromeDriver version mismatch
❌ Error: Chrome binary not found
```

**Solución:**

1. **Actualiza Google Chrome:**
   - Abre Chrome
   - Menú (⋮) → Ayuda → Información de Google Chrome
   - Espera a que se actualice automáticamente

2. **Descarga la última versión del ejecutable:**
   - La versión más reciente incluye ChromeDriver actualizado

3. **Verifica que Chrome esté instalado en la ruta predeterminada:**
   ```
   C:\Program Files\Google\Chrome\Application\chrome.exe
   ```

---

### Problema 3: La aplicación se cierra inmediatamente

**Síntomas:**
- La ventana aparece y desaparece en segundos
- No hay mensajes de error visibles

**Solución:**

1. **Ejecuta desde la terminal para ver errores:**
   ```cmd
   cd "ruta\donde\esta\la\app"
   SENA_Automation_App.exe
   ```

2. **Revisa el archivo de logs:**
   ```
   Logs/error_YYYYMMDD.log
   ```

3. **Verifica que todas las carpetas estén presentes:**
   - `funciones_formularios/`
   - `funciones_loggs/`
   - `perfilesOcupacionales/`
   - `URLS/`

4. **Reinstala desde cero:**
   - Descarga nuevamente el `.zip`
   - Extrae en una carpeta nueva
   - No copies archivos viejos

---

### Problema 4: Error al leer el archivo Excel

**Síntomas:**
```
❌ Error: No se pudo leer el archivo Excel
❌ Error: Columna 'X' no encontrada
```

**Solución:**

1. **Usa la plantilla proporcionada:**
   - Descarga: `plantilla_colocados.xlsx`
   - Copia tus datos manteniendo los nombres de columnas

2. **Verifica el formato del archivo:**
   - Debe ser `.xls` o `.xlsx`
   - No uses `.csv` o `.txt`

3. **Cierra el archivo en Excel antes de procesarlo:**
   - Excel bloquea archivos abiertos

4. **Verifica que no haya espacios extra en los nombres de columnas:**
   ```
   ✅ "Tipo de Documento"
   ❌ "Tipo de Documento " (espacio al final)
   ```

---

### Problema 5: Credenciales no se guardan

**Síntomas:**
- La app pide credenciales cada vez que se ejecuta
- Error: "No se pudo guardar configuración"

**Solución:**

1. **Verifica permisos de escritura:**
   - Clic derecho en la carpeta de la app
   - Propiedades → Seguridad
   - Tu usuario debe tener permisos de "Modificar"

2. **Ejecuta como administrador (temporal):**
   - Clic derecho en el `.exe`
   - "Ejecutar como administrador"
   - Ingresa credenciales
   - Cierra y ejecuta normalmente

3. **Verifica que la carpeta `config/` se haya creado:**
   ```
   SENA_Automation_App/
   └── config/
       └── credentials.json
   ```

---

### Problema 6: La página de APE no carga

**Síntomas:**
- Timeout después de varios segundos
- "No se pudo conectar con APE"

**Solución:**

1. **Verifica tu conexión a internet**

2. **Accede manualmente a APE:**
   - Abre Chrome y ve a la URL de APE
   - Verifica que funcione normalmente

3. **Revisa si APE está en mantenimiento:**
   - Contacta soporte técnico de APE

4. **Desactiva VPN o proxy:**
   - Algunos proxies bloquean la automatización

---

## 🔄 Actualizaciones

### ¿Cómo actualizar a una nueva versión?

1. **Descarga la nueva versión:**
   - Ve a [Releases](https://github.com/SergioAndresG/inscritos_sena_ape/releases/latest)
   - Descarga el archivo más reciente

2. **Reemplaza el ejecutable:**
   - Borra el antiguo `SENA_Automation_App.exe`
   - Copia el nuevo en la misma carpeta

3. **⚠️ NO borres estas carpetas:**
   - `config/` (contiene tus credenciales)
   - `Logs/` (historial de ejecuciones)
   - `perfilesOcupacionales/` (mapeos personalizados)

### Historial de Versiones

#### v1.2.0 (Última estable)
- ✅ Mejoras en estabilidad de Selenium
- ✅ Validación mejorada de archivos Excel
- ✅ Interfaz gráfica optimizada
- 🐛 Corrección de errores menores

#### v1.1.0
- ✅ Soporte para más tipos de documento
- ✅ Sistema de logs mejorado
- ✅ Optimización de velocidad

#### v1.0.0 (Inicial)
- ✅ Primera versión funcional
- ✅ Automatización básica completa

---

## 👨‍💻 Para Desarrolladores

### Requisitos de Desarrollo

Si deseas modificar el código fuente o contribuir al proyecto:

```bash
# 1. Clonar repositorio
git clone https://github.com/SergioAndresG/inscritos_sena_ape.git
cd inscritos_sena_ape

# 2. Crear entorno virtual
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar desde código fuente
python gui.py
```

### Estructura del Proyecto

```
INSCRITOS_APE_CBA/
│
├── gui.py                          # Interfaz gráfica principal
├── automatizacion.py               # Lógica de automatización Selenium
├── requirements.txt                # Dependencias del proyecto
├── SENA_Automation_App.spec        # Configuración de PyInstaller
│
├── funciones_formularios/          # Módulos de llenado de formularios
│   ├── __init__.py
│   ├── formulario_datos_basicos.py
│   ├── formulario_empresa.py
│   └── formulario_fecha.py
│
├── funciones_loggs/                # Sistema de logging
│   ├── __init__.py
│   └── logger.py
│
├── perfilesOcupacionales/          # Mapeo de programas
│   └── mapeo_programas.json
│
├── Iconos/                         # Recursos gráficos
│   ├── app_icon.ico
│   └── logo.png
│
├── URLS/                           # Configuración de enlaces
│   └── urls_ape.json
│
├── Logs/                           # Logs generados (no en repo)
├── build/                          # Archivos de compilación (no en repo)
└── dist/                           # Ejecutables generados (no en repo)
```

### Compilar el Ejecutable

Para generar el archivo `.exe`:

```bash
# Instalar PyInstaller
pip install pyinstaller

# Compilar (usa el archivo .spec personalizado)
pyinstaller SENA_Automation_App.spec

# El ejecutable estará en:
# dist/SENA_Automation_App.exe
```

### Contribuir al Proyecto

1. **Fork** el repositorio
2. Crea una rama para tu feature:
   ```bash
   git checkout -b feature/nueva-funcionalidad
   ```
3. Commit tus cambios:
   ```bash
   git commit -m "Add: descripción de la funcionalidad"
   ```
4. Push a tu fork:
   ```bash
   git push origin feature/nueva-funcionalidad
   ```
5. Abre un **Pull Request**

### Estilo de Código

- Usa **PEP 8** para Python
- Documenta funciones con **docstrings**
- Comenta código complejo
- Usa nombres descriptivos de variables

---

## 📧 Contacto y Soporte

### ¿Necesitas ayuda?

- 🐛 **Reportar bugs**: [Issues del repositorio](https://github.com/SergioAndresG/inscritos_sena_ape/issues)
- 💡 **Sugerencias**: [Discussions](https://github.com/SergioAndresG/inscritos_sena_ape/discussions)
- 📧 **Contacto directo**: [tu-email@sena.edu.co]

### Documentación Adicional

- 📖 [Manual de Usuario Completo](docs/MANUAL_USUARIO.pdf) (PDF)
- 🎥 [Video Tutorial](https://youtube.com/...) (Próximamente)
- 📝 [FAQ Extendido](docs/FAQ.md)

---

## 📊 Estadísticas de Impacto

Desde su implementación:

| Métrica | Valor |
|---------|-------|
| ⏱️ **Tiempo ahorrado** | ~85% de reducción |
| 📊 **Registros procesados** | +500 aprendices |
| ✅ **Tasa de éxito** | 98% |
| 👥 **Usuarios activos** | 15 funcionarios |

---

<p align="center">
  <strong>Desarrollado con ❤️ para optimizar el trabajo del SENA</strong>
</p>

<p align="center">
  <sub>Esta herramienta fue creada para ahorrar tiempo y reducir errores en procesos administrativos repetitivos</sub>
</p>

<p align="center">
  <a href="#-tabla-de-contenidos">⬆️ Volver arriba</a>
</p>
