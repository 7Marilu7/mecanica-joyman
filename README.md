# Mecánica JOYMAN

Sistema web de gestión de citas para un taller mecánico, desarrollado como proyecto de portafolio.

## Descripción

Mecánica JOYMAN es una aplicación web diseñada para facilitar la organización de clientes, servicios y citas de un taller mecánico. Cuenta con una interfaz personalizada con una temática automotriz en negro y amarillo.

## Funcionalidades

- Registro y edición de clientes.
- Registro y edición de servicios mecánicos.
- Gestión de citas con validación de fechas y horarios.
- Prevención de citas duplicadas en la misma fecha y hora.
- Edición y cancelación de citas.
- Registro de citas completadas e historial.
- Filtro del historial por fecha.
- Exportación del historial de citas a Excel.
- Panel principal con un resumen de la actividad del día.

## Tecnologías utilizadas

- Python
- Flask
- SQLite
- HTML5
- CSS3
- openpyxl

## Estructura del proyecto

```text
Mecanica_JOYMAN/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── static/
└── templates/
```

## Instalación local

1. Instala Python 3.
2. Descarga o clona este repositorio.
3. Crea un entorno virtual:

   ```bash
   python -m venv .venv
   ```

4. Activa el entorno virtual en Windows:

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

5. Instala las dependencias:

   ```bash
   python -m pip install -r requirements.txt
   ```

6. Ejecuta la aplicación:

   ```bash
   python app.py
   ```

7. Abre en tu navegador la dirección local que muestra Flask, normalmente `http://127.0.0.1:5000`.

La aplicación inicializa las tablas y los servicios de ejemplo al arrancar por medio de su función de inicialización.

## Base de datos

Por seguridad, la base de datos local no se incluye en este repositorio. La aplicación está preparada para trabajar con una base de datos SQLite ubicada junto al archivo `app.py`.

## Estado del proyecto

Este repositorio tiene fines de demostración y portafolio. No representa una instalación de producción del taller y no debe utilizarse con datos reales sin las medidas de seguridad, configuración y respaldos correspondientes.

## Autoría

Proyecto personal desarrollado para practicar y demostrar habilidades de desarrollo web con Python y Flask.
