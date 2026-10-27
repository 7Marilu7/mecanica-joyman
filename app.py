import sqlite3
from pathlib import Path
from datetime import date, time
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
import secrets

from flask import Flask, render_template, request, redirect, url_for, flash, send_file

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)

BASE_DIR = Path(__file__).resolve().parent
DATABASE = BASE_DIR / "database.db"


def conectar_db():
    """Conecta con la base de datos SQLite."""
    conexion = sqlite3.connect(DATABASE, timeout=10)
    conexion.row_factory = sqlite3.Row
    conexion.execute("PRAGMA foreign_keys = ON")
    return conexion


def inicializar_db():
    """Crea las tablas de la aplicación."""
    with conectar_db() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                telefono TEXT NOT NULL,
                correo TEXT,
                vehiculo TEXT
            )
        """)
        db.execute("""
            CREATE TABLE IF NOT EXISTS servicios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL UNIQUE,
                descripcion TEXT,
                precio REAL NOT NULL DEFAULT 0,
                activo INTEGER NOT NULL DEFAULT 1
            )
        """)

        db.execute("""
            CREATE TABLE IF NOT EXISTS citas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id INTEGER NOT NULL,
                servicio_id INTEGER NOT NULL,
                fecha TEXT NOT NULL,
                hora TEXT NOT NULL,
                estado TEXT NOT NULL DEFAULT 'confirmada',
                notas TEXT,

                FOREIGN KEY (cliente_id)
                    REFERENCES clientes(id),

                FOREIGN KEY (servicio_id)
                    REFERENCES servicios(id),

                CHECK (
                    estado IN (
                        'confirmada',
                        'completada',
                        'cancelada'
                    )
                )
            )
        """)

        # Impedir dos citas activas con la misma fecha y hora.
        # Las citas canceladas no bloquean el horario.
        db.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS
            evitar_citas_duplicadas
            ON citas(fecha, hora)
            WHERE estado != 'cancelada'
        """)

        servicios_iniciales = [
            (
                "Cambio de aceite",
                "Cambio de aceite y revisión básica",
                0
            ),
            (
                "Afinación",
                "Servicio de afinación del vehículo",
                0
            ),
            (
                "Revisión de frenos",
                "Inspección del sistema de frenos",
                0
            ),
            (
                "Diagnóstico mecánico",
                "Revisión para detectar fallas",
                0
            )
        ]

        db.executemany("""
            INSERT OR IGNORE INTO servicios
                (nombre, descripcion, precio)
            VALUES (?, ?, ?)
        """, servicios_iniciales)

    print("Base de datos inicializada correctamente.")



@app.route("/")
def inicio():
    hoy = date.today().isoformat()

    with conectar_db() as db:
        # Total de clientes registrados.
        total_clientes = db.execute("""
            SELECT COUNT(*)
            FROM clientes
        """).fetchone()[0]

        # Total de servicios activos.
        total_servicios = db.execute("""
            SELECT COUNT(*)
            FROM servicios
            WHERE activo = 1
        """).fetchone()[0]

        # Citas de hoy, excluyendo las canceladas.
        total_citas_hoy = db.execute("""
            SELECT COUNT(*)
            FROM citas
            WHERE fecha = ?
              AND estado != 'cancelada'
        """, (hoy,)).fetchone()[0]

        # Lista de las citas de hoy.
        citas_hoy = db.execute("""
            SELECT
                c.hora,
                c.estado,
                cl.nombre AS cliente,
                cl.vehiculo AS vehiculo,
                s.nombre AS servicio
            FROM citas c
            JOIN clientes cl ON cl.id = c.cliente_id
            JOIN servicios s ON s.id = c.servicio_id
            WHERE c.fecha = ?
              AND c.estado != 'cancelada'
            ORDER BY c.hora
        """, (hoy,)).fetchall()

    return render_template(
        "index.html",
        hoy=hoy,
        total_clientes=total_clientes,
        total_servicios=total_servicios,
        total_citas_hoy=total_citas_hoy,
        citas_hoy=citas_hoy
    )



@app.route("/clientes", methods=["GET", "POST"])
def clientes():
    if request.method == "POST":

        # Obtener los datos del formulario
        nombre = request.form.get("nombre", "").strip()
        telefono = request.form.get("telefono", "").strip()
        correo = request.form.get("correo", "").strip()
        vehiculo = request.form.get("vehiculo", "").strip()

        # Validar los campos obligatorios
        if not nombre or not telefono:
            flash(
                "El nombre y el teléfono son obligatorios.",
                "error"
            )
            return redirect(url_for("clientes"))

        if len(nombre) > 120 or len(telefono) > 30:
            flash(
                "El nombre o el teléfono son demasiado largos.",
                "error"
            )
            return redirect(url_for("clientes"))

        if len(correo) > 254 or len(vehiculo) > 150:
            flash(
                "El correo o los datos del vehículo son demasiado largos.",
                "error"
            )
            return redirect(url_for("clientes"))

        try:
            with conectar_db() as db:
                db.execute("""
                    INSERT INTO clientes
                        (nombre, telefono, correo, vehiculo)
                    VALUES (?, ?, ?, ?)
                """, (
                    nombre,
                    telefono,
                    correo or None,
                    vehiculo or None
                ))

            flash("Cliente registrado correctamente.", "exito")

        except sqlite3.Error:
            app.logger.exception(
                "Error al registrar el cliente."
            )
            flash("No se pudo registrar el cliente.", "error")

        return redirect(url_for("clientes"))

    with conectar_db() as db:
        lista_clientes = db.execute("""
            SELECT *
            FROM clientes
            ORDER BY nombre COLLATE NOCASE
        """).fetchall()

    return render_template(
        "clientes.html",
        clientes=lista_clientes
    )

@app.route("/clientes/editar/<int:cliente_id>", methods=["GET", "POST"])
def editar_cliente(cliente_id):
    with conectar_db() as db:
        cliente = db.execute("""
            SELECT id, nombre, telefono, correo, vehiculo
            FROM clientes
            WHERE id = ?
        """, (cliente_id,)).fetchone()

        if cliente is None:
            flash("El cliente no existe.", "error")
            return redirect(url_for("clientes"))

        if request.method == "POST":
            nombre = request.form.get("nombre", "").strip()
            telefono = request.form.get("telefono", "").strip()
            correo = request.form.get("correo", "").strip()
            vehiculo = request.form.get("vehiculo", "").strip()

            if not nombre or not telefono:
                flash("El nombre y el teléfono son obligatorios.", "error")
                return redirect(url_for("editar_cliente", cliente_id=cliente_id))

            if len(nombre) > 120 or len(telefono) > 30:
                flash("El nombre o el teléfono son demasiado largos.", "error")
                return redirect(url_for("editar_cliente", cliente_id=cliente_id))

            if len(correo) > 254 or len(vehiculo) > 150:
                flash("El correo o los datos del vehículo son demasiado largos.", "error")
                return redirect(url_for("editar_cliente", cliente_id=cliente_id))

            try:
                db.execute("""
                    UPDATE clientes
                    SET nombre = ?,
                        telefono = ?,
                        correo = ?,
                        vehiculo = ?
                    WHERE id = ?
                """, (
                    nombre,
                    telefono,
                    correo or None,
                    vehiculo or None,
                    cliente_id
                ))

                flash("Cliente actualizado correctamente.", "exito")
                return redirect(url_for("clientes"))

            except sqlite3.Error:
                app.logger.exception("Error al actualizar el cliente.")
                flash("No se pudo actualizar el cliente.", "error")
                return redirect(url_for("editar_cliente", cliente_id=cliente_id))

    return render_template("editar_cliente.html", cliente=cliente)

@app.route("/servicios", methods=["GET", "POST"])
def servicios():
    if request.method == "POST":

        nombre = request.form.get("nombre", "").strip()
        descripcion = request.form.get("descripcion", "").strip()
        precio_texto = request.form.get("precio", "").strip()

        # Validar los campos obligatorios
        if not nombre or not precio_texto:
            flash(
                "Completa todos los campos obligatorios.",
                "error"
            )
            return redirect(url_for("servicios"))

        # Validar el precio
        try:
            precio = float(precio_texto)

            if not 0 <= precio <= 10000000:
                raise ValueError

        except ValueError:
            flash("Introduce un precio válido.", "error")
            return redirect(url_for("servicios"))

        if len(nombre) > 100 or len(descripcion) > 500:
            flash(
                "El nombre o la descripción son demasiado largos.",
                "error"
            )
            return redirect(url_for("servicios"))

        try:
            with conectar_db() as db:
                db.execute("""
                    INSERT INTO servicios
                        (nombre, descripcion, precio)
                    VALUES (?, ?, ?)
                """, (
                    nombre,
                    descripcion or None,
                    precio
                ))

            flash("Servicio registrado correctamente.", "exito")

        except sqlite3.IntegrityError:
            flash(
                "Ya existe un servicio con ese nombre.",
                "error"
            )

        except sqlite3.Error:
            app.logger.exception(
                "Error al registrar el servicio."
            )
            flash("No se pudo registrar el servicio.", "error")

        return redirect(url_for("servicios"))

    with conectar_db() as db:
        lista_servicios = db.execute("""
            SELECT *
            FROM servicios
            WHERE activo = 1
            ORDER BY nombre COLLATE NOCASE
        """).fetchall()

    return render_template(
        "servicios.html",
        servicios=lista_servicios
    )

@app.route("/servicios/editar/<int:servicio_id>", methods=["GET", "POST"])
def editar_servicio(servicio_id):
    with conectar_db() as db:
        servicio = db.execute("""
            SELECT id, nombre, descripcion, precio, activo
            FROM servicios
            WHERE id = ?
        """, (servicio_id,)).fetchone()

        if servicio is None or servicio["activo"] != 1:
            flash("El servicio no existe o está desactivado.", "error")
            return redirect(url_for("servicios"))

        if request.method == "POST":
            nombre = request.form.get("nombre", "").strip()
            descripcion = request.form.get("descripcion", "").strip()
            precio_texto = request.form.get("precio", "").strip()

            if not nombre or not precio_texto:
                flash("El nombre y el precio son obligatorios.", "error")
                return redirect(
                    url_for("editar_servicio", servicio_id=servicio_id)
                )

            if len(nombre) > 100 or len(descripcion) > 500:
                flash("El nombre o la descripción son demasiado largos.", "error")
                return redirect(
                    url_for("editar_servicio", servicio_id=servicio_id)
                )

            try:
                precio = float(precio_texto)
                if not 0 <= precio <= 10000000:
                    raise ValueError
            except ValueError:
                flash("Introduce un precio válido.", "error")
                return redirect(
                    url_for("editar_servicio", servicio_id=servicio_id)
                )

            try:
                db.execute("""
                    UPDATE servicios
                    SET nombre = ?,
                        descripcion = ?,
                        precio = ?
                    WHERE id = ?
                """, (
                    nombre,
                    descripcion or None,
                    precio,
                    servicio_id
                ))

                flash("Servicio actualizado correctamente.", "exito")
                return redirect(url_for("servicios"))

            except sqlite3.IntegrityError:
                flash("Ya existe otro servicio con ese nombre.", "error")
                return redirect(
                    url_for("editar_servicio", servicio_id=servicio_id)
                )

            except sqlite3.Error:
                app.logger.exception("Error al actualizar el servicio.")
                flash("No se pudo actualizar el servicio.", "error")
                return redirect(
                    url_for("editar_servicio", servicio_id=servicio_id)
                )

    return render_template("editar_servicio.html", servicio=servicio)


@app.route("/servicios/eliminar/<int:servicio_id>", methods=["POST"])
def eliminar_servicio(servicio_id):
    try:
        with conectar_db() as db:
            servicio = db.execute("""
                SELECT id, nombre, activo
                FROM servicios
                WHERE id = ?
            """, (servicio_id,)).fetchone()

            if servicio is None:
                flash("El servicio no existe.", "error")
                return redirect(url_for("servicios"))

            if servicio["activo"] != 1:
                flash("Este servicio ya está desactivado.", "error")
                return redirect(url_for("servicios"))

            tiene_citas = db.execute("""
                SELECT 1
                FROM citas
                WHERE servicio_id = ?
                LIMIT 1
            """, (servicio_id,)).fetchone()

            if tiene_citas:
                db.execute("""
                    UPDATE servicios
                    SET activo = 0
                    WHERE id = ?
                """, (servicio_id,))

                flash(
                    "El servicio tenía citas relacionadas. "
                    "Se desactivó para conservar el historial.",
                    "exito"
                )
            else:
                db.execute("""
                    DELETE FROM servicios
                    WHERE id = ?
                """, (servicio_id,))

                flash("Servicio eliminado correctamente.", "exito")

    except sqlite3.Error:
        app.logger.exception("Error al eliminar el servicio.")
        flash("No se pudo eliminar el servicio.", "error")

    return redirect(url_for("servicios"))

@app.route("/citas", methods=["GET", "POST"])
def citas():

    if request.method == "POST":

        cliente_id = request.form.get(
            "cliente_id", ""
        ).strip()

        servicio_id = request.form.get(
            "servicio_id", ""
        ).strip()

        fecha = request.form.get("fecha", "").strip()
        hora = request.form.get("hora", "").strip()
        notas = request.form.get("notas", "").strip()

        # Validar los campos obligatorios
        if not all([cliente_id, servicio_id, fecha, hora]):
            flash(
                "Completa todos los campos obligatorios.",
                "error"
            )
            return redirect(url_for("citas"))

        if len(notas) > 1000:
            flash(
                "Las notas no pueden superar 1000 caracteres.",
                "error"
            )
            return redirect(url_for("citas"))

        try:
            cliente_id = int(cliente_id)
            servicio_id = int(servicio_id)

            fecha_obj = date.fromisoformat(fecha)
            hora_obj = time.fromisoformat(hora)

        except (ValueError, TypeError):
            flash(
                "Los datos de la cita no son válidos.",
                "error"
            )
            return redirect(url_for("citas"))

        # Evitar fechas pasadas
        if fecha_obj < date.today():
            flash(
                "No puedes reservar una fecha pasada.",
                "error"
            )
            return redirect(url_for("citas"))

        # Horario de atención: 07:00 a 18:00.
        # Se permiten citas desde las 07:00 hasta las 17:59.
        if not time(7, 0) <= hora_obj < time(18, 0):
            flash(
                "El horario de atención es de 07:00 a 18:00.",
                "error"
            )
            return redirect(url_for("citas"))

        # Normalizar la hora para guardar el formato HH:MM.
        hora = hora_obj.strftime("%H:%M")

        try:
            with conectar_db() as db:

                # Comprobar que el cliente existe
                cliente = db.execute("""
                    SELECT id
                    FROM clientes
                    WHERE id = ?
                """, (cliente_id,)).fetchone()

                # Comprobar que el servicio existe y está activo.
                servicio = db.execute("""
                    SELECT id
                    FROM servicios
                    WHERE id = ? AND activo = 1
                """, (servicio_id,)).fetchone()

                if cliente is None:
                    flash(
                        "El cliente seleccionado no existe.",
                        "error"
                    )
                    return redirect(url_for("citas"))

                if servicio is None:
                    flash(
                        "El servicio seleccionado no existe.",
                        "error"
                    )
                    return redirect(url_for("citas"))
                
                # Solo se bloquea la misma fecha y hora exactas  
                cita_existente = db.execute("""
                    SELECT id
                    FROM citas
                    WHERE fecha = ?
                      AND hora = ?
                      AND estado != 'cancelada'
                """, (fecha, hora)).fetchone()

                if cita_existente:
                    flash(
                        "Ya existe una cita registrada a esa hora. "
                        "Selecciona otra hora.",
                        "error"
                    )
                    return redirect(url_for("citas"))

                # Guardar la cita
                db.execute("""
                    INSERT INTO citas (
                        cliente_id,
                        servicio_id,
                        fecha,
                        hora,
                        estado,
                        notas
                    )
                    VALUES (?, ?, ?, ?, 'confirmada', ?)
                """, (
                    cliente_id,
                    servicio_id,
                    fecha,
                    hora,
                    notas or None
                ))

            flash(
                "Cita registrada correctamente.",
                "exito"
            )

        except sqlite3.IntegrityError:
            flash(
                "Ese horario acaba de ocuparse. "
                "Selecciona otra hora.",
                "error"
            )

        except sqlite3.Error:
            app.logger.exception(
                "Error al registrar la cita."
            )
            flash(
                "No se pudo registrar la cita.",
                "error"
            )

        return redirect(url_for("citas"))
    
     # Obtener la fecha seleccionada en el filtro.
    fecha_filtro = request.args.get("fecha", "").strip()

    # Consultar clientes, servicios y citas para mostrarlos.
    with conectar_db() as db:

        lista_clientes = db.execute("""
            SELECT id, nombre, telefono, vehiculo
            FROM clientes
            ORDER BY nombre COLLATE NOCASE
        """).fetchall()

        lista_servicios = db.execute("""
            SELECT id, nombre, precio
            FROM servicios
            WHERE activo = 1
            ORDER BY nombre COLLATE NOCASE
        """).fetchall()

        
        lista_citas = db.execute("""
            SELECT
                c.id,
                c.fecha,
                c.hora,
                c.estado,
                c.notas,
                cl.nombre AS cliente,
                cl.telefono AS telefono,
                cl.vehiculo AS vehiculo,
                s.nombre AS servicio,
                s.precio AS precio
            FROM citas c
            JOIN clientes cl
                ON cl.id = c.cliente_id
            JOIN servicios s
                ON s.id = c.servicio_id
            WHERE c.estado = 'confirmada'
            AND (? = '' OR c.fecha = ?)
            ORDER BY c.fecha, c.hora
        """, (fecha_filtro, fecha_filtro)).fetchall()

    return render_template(
        "citas.html",
        clientes=lista_clientes,
        servicios=lista_servicios,
        citas=lista_citas,
        hoy=date.today().isoformat(),
        fecha_filtro=fecha_filtro
    )

@app.route("/historial-citas")
def historial_citas():
    fecha_filtro = request.args.get("fecha", "").strip()

    with conectar_db() as db:
        historial = db.execute("""
            SELECT
                c.id,
                c.fecha,
                c.hora,
                c.estado,
                c.notas,
                cl.nombre AS cliente,
                cl.telefono AS telefono,
                cl.vehiculo AS vehiculo,
                s.nombre AS servicio,
                s.precio AS precio
            FROM citas c
            JOIN clientes cl
                ON cl.id = c.cliente_id
            JOIN servicios s
                ON s.id = c.servicio_id
            WHERE c.estado = 'completada'
              AND (? = '' OR c.fecha = ?)
            ORDER BY c.fecha DESC, c.hora DESC
        """, (fecha_filtro, fecha_filtro)).fetchall()

    return render_template(
        "historial_citas.html",
        historial=historial,
        fecha_filtro=fecha_filtro
    )

@app.route("/historial-citas/excel")
def descargar_historial_excel():
    fecha_filtro = request.args.get("fecha", "").strip()

    with conectar_db() as db:
        historial = db.execute("""
            SELECT
                c.fecha,
                c.hora,
                cl.nombre AS cliente,
                cl.telefono AS telefono,
                cl.vehiculo AS vehiculo,
                s.nombre AS servicio,
                s.precio AS precio,
                c.estado
            FROM citas c
            JOIN clientes cl ON cl.id = c.cliente_id
            JOIN servicios s ON s.id = c.servicio_id
            WHERE c.estado = 'completada'
              AND (? = '' OR c.fecha = ?)
            ORDER BY c.fecha DESC, c.hora DESC
        """, (fecha_filtro, fecha_filtro)).fetchall()

    libro = Workbook()
    hoja = libro.active
    hoja.title = "Historial de citas"

    encabezados = [
        "Fecha",
        "Hora",
        "Cliente",
        "Teléfono",
        "Vehículo",
        "Servicio",
        "Precio (MXN)",
        "Estado"
    ]

    hoja.append(["MECÁNICA JOYMAN - HISTORIAL DE CITAS"])
    hoja.merge_cells("A1:H1")
    hoja["A1"].font = Font(
        bold=True,
        size=16,
        color="FFFFFF"
    )
    hoja["A1"].fill = PatternFill(
        "solid",
        fgColor="202020"
    )
    hoja["A1"].alignment = Alignment(
        horizontal="center"
    )
    hoja.row_dimensions[1].height = 30

    hoja.append([
        f"Filtro de fecha: {fecha_filtro}"
        if fecha_filtro else "Todas las fechas"
    ])
    hoja.merge_cells("A2:H2")

    hoja.append(encabezados)

    for celda in hoja[3]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="D4A017")
        celda.alignment = Alignment(horizontal="center")

    for cita in historial:
        hoja.append([
            cita["fecha"],
            cita["hora"],
            cita["cliente"],
            cita["telefono"],
            cita["vehiculo"] or "—",
            cita["servicio"],
            cita["precio"],
            cita["estado"].capitalize()
        ])

    # Formato de precios y filtros.
    for fila in range(4, hoja.max_row + 1):
        hoja.cell(row=fila, column=7).number_format = '"$"#,##0.00'

    if hoja.max_row >= 3:
        hoja.auto_filter.ref = f"A3:H{hoja.max_row}"

    hoja.freeze_panes = "A4"

    # Ajustar el ancho de las columnas.
    for columna in hoja.columns:
        letra = get_column_letter(columna[0].column)
        ancho = max(
            len(str(celda.value or ""))
            for celda in columna
        )
        hoja.column_dimensions[letra].width = min(ancho + 3, 35)

    archivo = BytesIO()
    libro.save(archivo)
    archivo.seek(0)

    nombre_archivo = (
        f"historial_citas_{fecha_filtro}.xlsx"
        if fecha_filtro
        else "historial_citas_completadas.xlsx"
    )

    return send_file(
        archivo,
        as_attachment=True,
        download_name=nombre_archivo,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
    
@app.route("/citas/editar/<int:cita_id>", methods=["GET", "POST"])
def editar_cita(cita_id):
    try:
        with conectar_db() as db:
            cita = db.execute("""
                SELECT id, cliente_id, servicio_id, fecha, hora, estado, notas
                FROM citas
                WHERE id = ?
            """, (cita_id,)).fetchone()

            if cita is None:
                flash("La cita no existe.", "error")
                return redirect(url_for("citas"))

            if cita["estado"] != "confirmada":
                flash(
                    "Solo puedes editar citas que estén confirmadas.",
                    "error"
                )
                return redirect(url_for("citas"))

            if request.method == "POST":
                cliente_id_texto = request.form.get("cliente_id", "").strip()
                servicio_id_texto = request.form.get("servicio_id", "").strip()
                fecha = request.form.get("fecha", "").strip()
                hora = request.form.get("hora", "").strip()
                notas = request.form.get("notas", "").strip()

                # Validar los campos obligatorios.
                if not all([
                    cliente_id_texto,
                    servicio_id_texto,
                    fecha,
                    hora
                ]):
                    flash("Completa todos los campos obligatorios.", "error")
                    return redirect(url_for("editar_cita", cita_id=cita_id))

                try:
                    cliente_id = int(cliente_id_texto)
                    servicio_id = int(servicio_id_texto)
                    fecha_obj = date.fromisoformat(fecha)
                    hora_obj = time.fromisoformat(hora)
                except ValueError:
                    flash("El cliente, servicio, fecha u hora no son válidos.", "error")
                    return redirect(url_for("editar_cita", cita_id=cita_id))

                # Validar fecha y horario del taller.
                if fecha_obj < date.today():
                    flash("No puedes programar una cita en una fecha pasada.", "error")
                    return redirect(url_for("editar_cita", cita_id=cita_id))

                if hora_obj < time(8, 0) or hora_obj >= time(18, 0):
                    flash("El horario permitido es de 08:00 a 17:59.", "error")
                    return redirect(url_for("editar_cita", cita_id=cita_id))

                # Comprobar que el cliente y el servicio existan.
                cliente = db.execute("""
                    SELECT id FROM clientes WHERE id = ?
                """, (cliente_id,)).fetchone()

                servicio = db.execute("""
                    SELECT id FROM servicios
                    WHERE id = ? AND (activo = 1 OR id = ?)
                """, (servicio_id, cita["servicio_id"])).fetchone()

                if cliente is None:
                    flash("El cliente seleccionado no existe.", "error")
                    return redirect(url_for("editar_cita", cita_id=cita_id))

                if servicio is None:
                    flash("El servicio seleccionado no está disponible.", "error")
                    return redirect(url_for("editar_cita", cita_id=cita_id))

                # Evitar otra cita activa en la misma fecha y hora.
                # Se excluye la propia cita que estamos editando.
                cita_existente = db.execute("""
                    SELECT id
                    FROM citas
                    WHERE fecha = ?
                      AND hora = ?
                      AND estado != 'cancelada'
                      AND id != ?
                """, (fecha, hora, cita_id)).fetchone()

                if cita_existente:
                    flash(
                        "Ya existe otra cita registrada en esa fecha y hora.",
                        "error"
                    )
                    return redirect(url_for("editar_cita", cita_id=cita_id))

                # Guardar los cambios.
                db.execute("""
                    UPDATE citas
                    SET cliente_id = ?,
                        servicio_id = ?,
                        fecha = ?,
                        hora = ?,
                        notas = ?
                    WHERE id = ?
                      AND estado = 'confirmada'
                """, (
                    cliente_id,
                    servicio_id,
                    fecha,
                    hora,
                    notas or None,
                    cita_id
                ))

                flash("La cita se actualizó correctamente.", "exito")
                return redirect(url_for("citas"))

            # Cargar clientes y servicios para el formulario.
            clientes = db.execute("""
                SELECT id, nombre, telefono, vehiculo
                FROM clientes
                ORDER BY nombre COLLATE NOCASE
            """).fetchall()

            servicios = db.execute("""
                SELECT id, nombre, precio
                FROM servicios
                WHERE activo = 1 OR id = ?
                ORDER BY nombre COLLATE NOCASE
            """, (cita["servicio_id"],)).fetchall()

            return render_template(
                "editar_cita.html",
                cita=cita,
                clientes=clientes,
                servicios=servicios,
                hoy=date.today().isoformat()
            )

    except sqlite3.IntegrityError:
        flash("Ese horario acaba de ocuparse. Elige otra hora.", "error")
        return redirect(url_for("editar_cita", cita_id=cita_id))

    except sqlite3.Error:
        app.logger.exception("Error al editar la cita.")
        flash("Ocurrió un error al actualizar la cita.", "error")
        return redirect(url_for("citas"))


@app.route("/citas/cancelar/<int:cita_id>", methods=["POST"])
def cancelar_cita(cita_id):

    try:
        with conectar_db() as db:

            cita = db.execute("""
                SELECT id, estado
                FROM citas
                WHERE id = ?
            """, (cita_id,)).fetchone()

            if cita is None:
                flash("La cita no existe.", "error")

            elif cita["estado"] == "cancelada":
                flash(
                    "Esta cita ya está cancelada.",
                    "error"
                )

            elif cita["estado"] == "completada":
                flash(
                    "No puedes cancelar una cita completada.",
                    "error"
                )

            else:
                db.execute("""
                    UPDATE citas
                    SET estado = 'cancelada'
                    WHERE id = ?
                """, (cita_id,))

                flash(
                    "Cita cancelada. El horario está disponible.",
                    "exito"
                )

    except sqlite3.Error:
        app.logger.exception(
            "Error al cancelar la cita."
        )
        flash(
            "No se pudo cancelar la cita.",
            "error"
        )

    return redirect(url_for("citas"))

@app.route("/citas/completar/<int:cita_id>", methods=["POST"])
def completar_cita(cita_id):
    with conectar_db() as db:
        cita = db.execute(
            "SELECT id, estado FROM citas WHERE id = ?",
            (cita_id,)
        ).fetchone()

        if cita is None:
            flash("La cita no existe.", "error")
        elif cita["estado"] != "confirmada":
            flash("Solo puedes completar citas confirmadas.", "error")
        else:
            db.execute(
                "UPDATE citas SET estado = 'completada' WHERE id = ?",
                (cita_id,)
            )
            flash("¡Cita marcada como completada!", "exito")

    return redirect(url_for("citas"))


@app.route("/citas/eliminar/<int:cita_id>", methods=["POST"])
def eliminar_cita(cita_id):
    try:
        with conectar_db() as db:
            cita = db.execute(
                "SELECT id FROM citas WHERE id = ?",
                (cita_id,)
            ).fetchone()

            if cita is None:
                flash("La cita no existe.", "error")
                return redirect(url_for("citas"))

            db.execute(
                "DELETE FROM citas WHERE id = ?",
                (cita_id,)
            )

            flash("Cita eliminada correctamente.", "exito")

    except sqlite3.Error:
        app.logger.exception("Error al eliminar la cita.")
        flash("No se pudo eliminar la cita.", "error")

    return redirect(url_for("citas"))



if __name__ == "__main__":
    inicializar_db()
    app.run(debug=True)