#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
construir_xlsm.py
Reconstruye el libro Excel TiendaFutbol.xlsm con diseño visual ejecutivo.
Requisitos: Windows, Excel de escritorio, pywin32 (pip install pywin32) y activar en Excel
  Archivo > Opciones > Centro de confianza > Configuración de macros >
  "Confiar en el acceso al modelo de objetos de proyectos de VBA".
Resultado: al abrir el libro solo se ve la hoja "Inicio" (logo + botones ABRIR TIENDA / SALIR);
las demás hojas quedan muy ocultas y todo el trabajo se hace en formularios.

Pasos:
  1. Crea las hojas del sistema (incluyendo 'Catalogo' y 'PlantillaBoleta') y
     formatea la hoja 'Configuracion' con los rangos nombrados:
       - RutaBaseDatos (celda A1)
       - RutaBoletas   (celda A2)
       - CorreoPrueba  (celda A3)
       - GmailUsuario  (celda A4)
       - GmailClaveApp (celda A5 -> vacía por seguridad en Git)
       - RutaImagenes  (celda A6 -> carpeta imagenes_productos/)
  2. Normaliza el encabezado 'Attribute VB_Name' en la línea 1 de los archivos .bas.
  3. Importa los módulos .bas de /excel/codigo_vba/.
  4. Construye programáticamente los UserForms de /excel/codigo_vba/formularios/
     (incluyendo frmCatalogo con MSForms.Image) aplicando diseño ejecutivo moderno.
  5. Vuelca inicialmente el catálogo, inventario, ventas, clientes y proveedores
     a las hojas de Excel para que el libro nunca quede con hojas vacías.
"""

import glob
import os
import re
import sys
import traceback
import win32com.client as win32

MODULOS_2B1 = [
    "modConexion.bas",
    "modSeguridad.bas",
    "modValidaciones.bas",
    "modClientes.bas",
    "modUsuarios.bas",
]

MODULOS_ADICIONALES = [
    "modEstilo.bas",
    "modInventario.bas",
    "modVentas.bas",
    "modBoletas.bas",
    "modSesion.bas",
    "modInicio.bas",
]


def normalizar_bas(ruta_bas, nombre_forzado=None):
    """Asegura que 'Attribute VB_Name = "..."' esté en la línea 1 del archivo .bas."""
    nombre_mod = (
        nombre_forzado or os.path.splitext(os.path.basename(ruta_bas))[0]
    )
    with open(ruta_bas, "r", encoding="utf-8", errors="replace") as f:
        lineas = f.read().splitlines()

    lineas_limpias = [
        ln
        for ln in lineas
        if not ln.strip().startswith("Attribute VB_Name")
        and not ln.strip().startswith("\ufeffAttribute VB_Name")
    ]
    encabezado = f'Attribute VB_Name = "{nombre_mod}"'
    contenido = "\r\n".join([encabezado] + lineas_limpias) + "\r\n"
    with open(ruta_bas, "w", encoding="utf-8", newline="") as f:
        f.write(contenido)


def parsear_frm_texto(ruta_frm):
    """Lee un archivo .frm con bloques MSForms en texto y devuelve (nombre, props_form, controles, codigo_vba)."""
    with open(ruta_frm, "r", encoding="utf-8", errors="replace") as f:
        texto = f.read()

    m_code = re.search(r"Attribute VB_Exposed\s*=\s*False\r?\n", texto)
    if not m_code:
        raise ValueError(f"No se encontró 'Attribute VB_Exposed = False' en {ruta_frm}")

    header = texto[: m_code.start()]
    codigo = texto[m_code.end() :].strip() + "\r\n"

    m_form = re.search(r"Begin\s+\{[0-9A-Fa-f\-]+\}\s+(\w+)", header)
    if not m_form:
        raise ValueError(f"No se encontró encabezado Begin {{GUID}} en {ruta_frm}")
    form_name = m_form.group(1)

    first_ctrl = header.find("Begin MSForms.")
    bloque_form = (
        header[m_form.end() : first_ctrl]
        if first_ctrl != -1
        else header[m_form.end() :]
    )
    props_form = dict(re.findall(r"^\s*(\w+)\s*=\s*(.+?)\s*$", bloque_form, re.M))

    controles = []
    for ctype, cname, cbody in re.findall(
        r"Begin\s+MSForms\.(\w+)\s+(\w+)\s*\r?\n(.*?)\r?\n\s*End", header, re.S
    ):
        cprops = dict(re.findall(r"^\s*(\w+)\s*=\s*(.+?)\s*$", cbody, re.M))
        controles.append((ctype, cname, cprops))

    return form_name, props_form, controles, codigo


def limpiar_valor_frm(val_str):
    v = val_str.strip()
    if v.startswith('"') and v.endswith('"'):
        return v[1:-1]
    if "'" in v:
        v = v.split("'")[0].strip()
    return v


def set_ctrl_prop(ctrl, nombre_prop, valor):
    try:
        setattr(ctrl, nombre_prop, valor)
        return
    except Exception:
        pass
    try:
        setattr(ctrl.Object, nombre_prop, valor)
    except Exception:
        pass


def construir_userform_desde_frm(vb_proj, ruta_frm):
    """Crea un UserForm real en Excel con estilo visual ejecutivo a partir de la definición .frm."""
    form_name, props_form, controles, codigo = parsear_frm_texto(ruta_frm)

    try:
        existente = vb_proj.VBComponents(form_name)
        vb_proj.VBComponents.Remove(existente)
    except Exception:
        pass

    comp = vb_proj.VBComponents.Add(3)
    comp.Name = form_name

    if "Caption" in props_form:
        comp.Properties("Caption").Value = limpiar_valor_frm(props_form["Caption"])
    if "ClientWidth" in props_form:
        w_pt = float(limpiar_valor_frm(props_form["ClientWidth"])) / 20.0
        comp.Properties("Width").Value = w_pt + 14
    if "ClientHeight" in props_form:
        h_pt = float(limpiar_valor_frm(props_form["ClientHeight"])) / 20.0
        comp.Properties("Height").Value = h_pt + 28
    if "StartUpPosition" in props_form:
        try:
            comp.Properties("StartUpPosition").Value = int(
                limpiar_valor_frm(props_form["StartUpPosition"])
            )
        except Exception:
            pass

    # Fondo gris pizarra claro (#F1F5F9 -> BGR 0xF9F5F1)
    try:
        comp.Properties("BackColor").Value = 0xF9F5F1
    except Exception:
        pass

    designer = comp.Designer
    for ctype, cname, cprops in controles:
        prog_id = f"Forms.{ctype}.1"
        ctrl = designer.Controls.Add(prog_id, cname, True)

        for geom_prop in ("Left", "Top", "Width", "Height"):
            if geom_prop in cprops:
                set_ctrl_prop(
                    ctrl,
                    geom_prop,
                    float(limpiar_valor_frm(cprops[geom_prop])) / 20.0,
                )

        try:
            ctrl.Font.Name = "Segoe UI"
            ctrl.Font.Size = 9
        except Exception:
            pass

        if "Caption" in cprops:
            set_ctrl_prop(ctrl, "Caption", limpiar_valor_frm(cprops["Caption"]))
        if "TabIndex" in cprops:
            set_ctrl_prop(ctrl, "TabIndex", int(limpiar_valor_frm(cprops["TabIndex"])))
        if "MaxLength" in cprops:
            set_ctrl_prop(
                ctrl, "MaxLength", int(limpiar_valor_frm(cprops["MaxLength"]))
            )
        if "PasswordChar" in cprops:
            pchar = limpiar_valor_frm(cprops["PasswordChar"])
            set_ctrl_prop(
                ctrl, "PasswordChar", chr(int(pchar)) if pchar.isdigit() else pchar
            )
        if "WordWrap" in cprops:
            set_ctrl_prop(
                ctrl,
                "WordWrap",
                limpiar_valor_frm(cprops["WordWrap"]) in ("-1", "True"),
            )
        if "Enabled" in cprops:
            set_ctrl_prop(
                ctrl,
                "Enabled",
                limpiar_valor_frm(cprops["Enabled"]) not in ("0", "False"),
            )
        if "ForeColor" in cprops:
            fc = limpiar_valor_frm(cprops["ForeColor"])
            if fc.startswith("&H") and fc.endswith("&"):
                set_ctrl_prop(ctrl, "ForeColor", int(fc[2:-1], 16))
        if "Style" in cprops:
            set_ctrl_prop(ctrl, "Style", int(limpiar_valor_frm(cprops["Style"])))
        if "Default" in cprops:
            is_def = limpiar_valor_frm(cprops["Default"]) in ("-1", "True")
            set_ctrl_prop(ctrl, "Default", is_def)
            if is_def and ctype == "CommandButton":
                try:
                    set_ctrl_prop(ctrl, "BackColor", 0x699605)  # Emerald #059669
                    set_ctrl_prop(ctrl, "ForeColor", 0xFFFFFF)
                    ctrl.Font.Bold = True
                except Exception:
                    pass
        if "Cancel" in cprops:
            set_ctrl_prop(
                ctrl,
                "Cancel",
                limpiar_valor_frm(cprops["Cancel"]) in ("-1", "True"),
            )

    code_mod = comp.CodeModule
    if code_mod.CountOfLines > 0:
        code_mod.DeleteLines(1, code_mod.CountOfLines)
    code_mod.AddFromString(codigo)

    return comp, form_name


def estilizar_hoja_configuracion(config_ws, wb, accdb_path, boletas_dir, imagenes_dir):
    """Configura los rangos nombrados en A1:A6 y agrega etiquetas descriptivas en B1:C6."""
    config_ws.Cells.Font.Name = "Segoe UI"
    config_ws.Cells.Font.Size = 10

    filas_cfg = [
        (
            "A1",
            "RutaBaseDatos",
            accdb_path,
            "Ruta archivo Access (.accdb)",
            "Base de datos principal del sistema",
        ),
        (
            "A2",
            "RutaBoletas",
            boletas_dir,
            "Carpeta Boletas PDF",
            "Destino donde se exportan <correlativo>.pdf",
        ),
        (
            "A3",
            "CorreoPrueba",
            "m.coroseo.ramos@gmail.com",
            "Correo de Prueba (Redirección)",
            "Recibe las boletas de clientes @example.com",
        ),
        (
            "A4",
            "GmailUsuario",
            "m.coroseo.ramos@gmail.com",
            "Cuenta Gmail Emisora (SMTP)",
            "Usuario para smtp.gmail.com:465 (SSL)",
        ),
        (
            "A5",
            "GmailClaveApp",
            "",
            "Contraseña de Aplicación Gmail",
            "App Password de 16 letras (dejar vacía en Git)",
        ),
        (
            "A6",
            "RutaImagenes",
            imagenes_dir,
            "Carpeta Imágenes de Productos",
            "Fotografías para el Catálogo Visual (frmCatalogo y hoja Catalogo)",
        ),
    ]

    for celda, nombre_rango, valor, etiqueta, nota in filas_cfg:
        r = config_ws.Range(celda)
        r.Value = valor
        wb.Names.Add(Name=nombre_rango, RefersTo=r)
        fila_idx = r.Row
        config_ws.Cells(fila_idx, 2).Value = etiqueta
        config_ws.Cells(fila_idx, 2).Font.Bold = True
        config_ws.Cells(fila_idx, 3).Value = nota
        config_ws.Cells(fila_idx, 3).Font.Color = 0x64748B

    config_ws.Columns("A:C").AutoFit


def main():
    solo_2b1 = "--solo-2b1" in sys.argv
    probar_2c = "--probar-2c" in sys.argv
    probar_3a = "--probar-3a" in sys.argv or "--probar-3b" in sys.argv

    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    excel_dir = os.path.join(base_dir, "excel")
    codigo_vba_dir = os.path.join(excel_dir, "codigo_vba")
    formularios_dir = os.path.join(codigo_vba_dir, "formularios")
    pruebas_dir = os.path.join(codigo_vba_dir, "pruebas_vba")
    xlsm_path = os.path.join(excel_dir, "TiendaFutbol.xlsm")
    accdb_path = os.path.join(base_dir, "access", "tienda_futbol.accdb")
    boletas_dir = os.path.join(base_dir, "boletas")
    imagenes_dir = os.path.join(base_dir, "imagenes_productos")
    os.makedirs(boletas_dir, exist_ok=True)
    os.makedirs(imagenes_dir, exist_ok=True)

    # Comprobaciones previas de recursos gráficos
    assets_dir = os.path.join(base_dir, "assets")
    faltan = [a for a in ("portada.jpg", "logo_form.jpg") if not os.path.exists(os.path.join(assets_dir, a))]
    if faltan:
        print("AVISO: faltan en assets/:", ", ".join(faltan), "-> ejecute: python scripts/generar_logo.py")
    n_img = len(glob.glob(os.path.join(imagenes_dir, "*.jpg")))
    if n_img < 60:
        print(f"AVISO: imagenes_productos/ tiene {n_img} fotos (se esperan 60) -> ejecute: python scripts/generar_imagenes.py")

    print("Iniciando Excel...")
    xl = None
    wb = None
    try:
        xl = win32.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = 1

        wb = xl.Workbooks.Add()
        sheet_names = [
            "Inicio",
            "Configuracion",
            "Catalogo",
            "Ventas",
            "Inventario",
            "Productos",
            "Clientes",
            "Proveedores",
            "Usuarios",
            "Boletas",
            "PlantillaBoleta",
        ]
        wb.Worksheets(1).Name = sheet_names[0]
        # La macro Workbook_Open/PrepararPortada necesita acceso al modelo VBA: comprobarlo ya
        try:
            _ = wb.VBProject.VBComponents.Count
        except Exception:
            raise SystemExit(
                "Excel bloquea el acceso al proyecto VBA. Active: Archivo > Opciones > Centro de confianza > "
                "Configuración del Centro de confianza > Configuración de macros > "
                "'Confiar en el acceso al modelo de objetos de proyectos de VBA' y vuelva a ejecutar."
            )
        for name in sheet_names[1:]:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = name

        config_ws = wb.Worksheets("Configuracion")
        estilizar_hoja_configuracion(config_ws, wb, accdb_path, boletas_dir, imagenes_dir)

        vb_proj = wb.VBProject
        modulos_a_importar = list(MODULOS_2B1)
        if not solo_2b1:
            for m in MODULOS_ADICIONALES:
                if os.path.exists(os.path.join(codigo_vba_dir, m)):
                    modulos_a_importar.append(m)

        for fname in modulos_a_importar:
            full_path = os.path.join(codigo_vba_dir, fname)
            if not os.path.exists(full_path):
                continue
            normalizar_bas(full_path)
            comp_name = os.path.splitext(fname)[0]
            try:
                vb_proj.VBComponents.Remove(vb_proj.VBComponents(comp_name))
            except Exception:
                pass
            comp = vb_proj.VBComponents.Import(full_path)
            if comp.Name != comp_name:
                comp.Name = comp_name
            print(f"  Importado módulo: {comp.Name}")

        if not solo_2b1 and os.path.isdir(formularios_dir):
            for ruta_frm in sorted(glob.glob(os.path.join(formularios_dir, "*.frm"))):
                comp, form_name = construir_userform_desde_frm(vb_proj, ruta_frm)
                print(
                    f"  UserForm construido en el libro: {form_name}"
                    f" ({comp.Designer.Controls.Count} controles)"
                )

        # Código del objeto ThisWorkbook (Workbook_Open, etc.)
        ruta_tw = os.path.join(codigo_vba_dir, "ThisWorkbook.vba")
        if not solo_2b1 and os.path.exists(ruta_tw):
            with open(ruta_tw, "r", encoding="utf-8", errors="replace") as f:
                codigo_tw = f.read().strip() + "\r\n"
            cm = vb_proj.VBComponents(wb.CodeName).CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(codigo_tw)
            print("  Código de ThisWorkbook inyectado")

        wb.SaveAs(xlsm_path, FileFormat=52)

        # Dibujar la portada (logo + botones) y ocultar las demás hojas.
        # Debe hacerse con el libro ya guardado para resolver ..\assets\
        if not solo_2b1:
            try:
                xl.Run(f"'{wb.Name}'!modInicio.PrepararPortada")
                xl.Run(f"'{wb.Name}'!modInicio.RestaurarEntorno")
                print("  Portada construida; hojas internas ocultas")
            except Exception as exc:
                print("  AVISO: no se pudo construir la portada ahora (se hará al abrir el libro):", exc)
            wb.Save()
        print(f"\nÉXITO: Libro guardado en {xlsm_path}")
    except Exception:
        traceback.print_exc()
        sys.exit(1)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
