<!-- Título Principal -->
<h1 align="center" style="font-size: 38px; font-weight: bold;">
✨ Automatización de Ingreso de Colocados en APE – SENA
</h1>

<!-- Subtítulo -->
<h3 align="center" style="color: #4E8DA6; font-weight: normal; margin-top: -10px;">
Automatización de carga masiva para la Agencia Pública de Empleo del SENA
</h3>

<!-- Badges -->
<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10-blue?logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/Selenium-Automation-43B02A?logo=selenium&logoColor=white" />
</p>

<br>

<!-- Separador bonito -->
<hr style="border: 1px solid #bfbfbf; width: 80%;">

<br>

<!-- Sección 1 -->
<h2 align="center">🎯 ¿Por qué usar esta herramienta?</h2>

<p align="center" style="max-width: 750px; margin: auto; font-size: 17px;">
El registro manual en la plataforma APE es lento, tedioso y propenso a errores.  
Esta aplicación convierte un proceso de <strong>horas en minutos</strong>, garantizando exactitud, trazabilidad y eficiencia.
</p>

<br>

<!-- Cards de Beneficios -->
<div align="center">
  <table>
    <tr>
      <td align="center" width="250">
        <h3>⏱️ Ahorro de Tiempo</h3>
        <p>Carga masiva desde Excel con un solo clic.</p>
      </td>
      <td align="center" width="250">
        <h3>🎯 Reducción de Errores</h3>
        <p>Automatiza el llenado de formularios.</p>
      </td>
      <td align="center" width="250">
        <h3>🧾 Trazabilidad</h3>
        <p>Logs detallados de cada acción.</p>
      </td>
    </tr>
  </table>
</div>

<br><br>

<!-- Sección 2 -->
<h2 align="center">⚙️ Características Técnicas</h2>

<div align="center">
  <p style="font-size: 16px; max-width: 700px;">
    <strong>•</strong> Interfaz gráfica intuitiva creada en Python (CustomTkinter).<br>
    <strong>•</strong> Procesamiento de archivos Excel (.xls) con validación.<br>
    <strong>•</strong> Automatización con Selenium para login, navegación y llenado.<br>
    <strong>•</strong> Registro de actividad mediante módulo logging.<br>
  </p>
</div>

<br><br>

<!-- Sección 3 -->
<h2 align="center">💻 Requisitos Previos</h2>

<div align="center">
  <table style="width: 80%; font-size: 16px;">
    <tr>
      <th>Recurso</th>
      <th>Descripción</th>
    </tr>
    <tr>
      <td><strong>Python 3.8+</strong></td>
      <td>Versión recomendada para ejecutar la aplicación.</td>
    </tr>
    <tr>
      <td><strong>WebDriver</strong></td>
      <td>ChromeDriver u otro según navegador. Versiones deben coincidir.</td>
    </tr>
    <tr>
      <td><strong>Dependencias</strong></td>
      <td>Incluidas en <code>requeriments.txt</code>.</td>
    </tr>
  </table>
</div>

<br><br>




<br>

<!-- Sección 4 -->
<h2 align="center">🔐 Configuración Necesaria Antes de Ejecutar</h2>

<div align="center">
  <p style="max-width: 750px; font-size: 16px;">
    Antes de ejecutar la aplicación por primera vez, es necesario ingresar la información requerida 
    para que la automatización pueda iniciar sesión y seleccionar correctamente el perfil ocupacional.
  </p>

  <table style="width: 80%; font-size: 16px; margin-top: 20px;">
    <tr>
      <th>Configuración</th>
      <th>Descripción</th>
    </tr>
    <tr>
      <td><strong>Credenciales APE</strong></td>
      <td>Debes ingresar el usuario y contraseña para acceder a la plataforma APE.</td>
    </tr>
    <tr>
      <td><strong>Perfil Ocupacional</strong></td>
      <td>
        Si no existe un perfil relacionado con el nombre del programa, la aplicación solicitará uno 
        antes de continuar. Este dato es obligatorio.
      </td>
    </tr>
  </table>
</div>

<br>

<!-- Sección 5 -->
<h2 align="center">🚀 Instalación</h2>

<p align="center">
  <strong>1️⃣ Clonar repositorio</strong>
</p>

```bash
git clone https://github.com/SergioAndresG/inscritos_sena_ape.git
```


