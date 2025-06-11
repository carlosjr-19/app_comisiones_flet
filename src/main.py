import flet as ft
import pandas as pd
import os
from service import comisiones_act, comisiones_rec
from service import env

def main(page: ft.Page):    
    # Variables de estado para archivos seleccionados
    archivo_1_path = ""
    archivo_2_path = ""

    #Marcas
    mvnos = env.MVNOS

################################## FUNCIONES #########################################################
    def on_file_picker_1_result(e):
        nonlocal archivo_1_path
        if e.files:
            archivo_1_path = e.files[0].path
            resumen_text.value += (f"\n📑Archivo 1: {archivo_1_path}")
            page.update()

    def on_file_picker_2_result(e):
        nonlocal archivo_2_path
        if e.files:
            archivo_2_path = e.files[0].path
            resumen_text.value += (f"\n📑Archivo 2: {archivo_2_path}")
            page.update()
    
    def procesar_archivos(e):
        if not archivo_1_path or not archivo_2_path:
            resumen_text.value = "Debes seleccionar ambos archivos"
            page.update()
            return

        try:
            loader.visible = True 
            page.update()

            df1 = pd.read_csv(archivo_1_path)
            df2 = pd.read_excel(archivo_2_path)

            comision = input_comision.value
            sales_p = select_pago_sales.value
            marca = select_marca.value
            proceso = select_proceso.value
            fecha = input_fecha.value

            if proceso == 'Activación':
                print('Bloque activación')
                df2 = df2.query(f'mvno_name == "{marca}"' )
                print("Reporte Marca: ", len(df1))
                print("Reporte General: ", len(df2))

                resumen_text.value = (
                    f"\n##############################################################\n\n"
                )

                duplicados, csv_sin_duplicados = comisiones_act.limpiar_duplicados(df1, df2)     

                if duplicados.empty:
                    resumen_text.value = "No hay duplicados\n"
                    xlsx, precios_iguales = comisiones_act.procesar_comisiones(df1, sales_p, comision, fecha)
                else:
                    df1 = csv_sin_duplicados
                    resumen_text.value = "Si hay duplicados y fueron limpiados\n"
                    xlsx, precios_iguales = comisiones_act.procesar_comisiones(df1, sales_p, comision, fecha)

                comisiones_act.estilos_excel(xlsx, marca, precios_iguales, fecha)

                resumen_text.value += (
                    #f"📑 Archivo 1: {df1}\n"
                    #f"📑 Archivo 2: {df2}\n\n"
                    f"✔️ Total líneas CSV {marca}: {len(df1)}\n"
                    f"✔️ Total líneas en excel general: {len(df2)}\n"
                    f"❌ Números duplicados: {duplicados[['msisdn']]}\n"
                    f"➡️ Comisión: {comision}%\n"
                    f"➡️ Marca: {marca}\n"
                    f"➡️ Se paga por sales: {sales_p}\n"
                    f"➡️ Proceso: {proceso}"
                )

            elif proceso == 'Recarga':
                print('Bloque recarga')
                df2 = df2.query(f'name == "{marca}"')
                print("cantidad antes de eliminar lineas con 1 en reporte marca:", len(df1))

                resumen_text.value += (
                    f"\n##############################################################\n\n"
                    f"✔️ cantidad antes de eliminar lineas con 1 en reporte marca: {len(df1)}\n"
                )

                print("cantidad en el reporte general:", len(df2))

                # Filtrar y eliminar los msisdn que comienzan con 1 en csv
                df1 = df1[~df1['msisdn'].astype(str).str.startswith('1')]

                # Confirmar nuevo total
                print(f"Total en CSV {marca} después de eliminar los que empiezan con 1: {len(df1)}")

                csv, diferencias = comisiones_rec.limpiar_archivo(df1, df2)
                xlsx, precios_iguales = comisiones_rec.procesar_comisiones(csv, sales_p, comision, fecha)
                comisiones_rec.estilos_excel(xlsx, marca, precios_iguales, fecha)

                resumen_text.value += (
                    #f"📑 Archivo 1: {df1}\n"
                    #f"📑 Archivo 2: {df2}\n\n"
                    #f"✔️ cantidad antes de eliminar lineas con 1 en reporte marca: {len(csv)}\n"
                    f"✔️ Total en CSV Marca después de eliminar los que empiezan con 1: {len(df1)}\n"
                    f"✔️ Total líneas CSV 1: {len(df1)}\n"
                    f"✔️ Total líneas CSV 2: {len(df2)}\n"
                    f"✔️ Diferencias en número de recargas por línea:\n {diferencias}\n"
                    f"➡️ Comisión: {comision}%\n"
                    f"➡️ Marca: {marca}\n"
                    f"➡️ Se paga por sales: {sales_p}\n"
                    f"➡️ Proceso: {proceso}"
                )

            else:
                resumen_text.value += f"❌ Selecciona un proceso\n"

        except Exception as ex:
                resumen_text.value += f"❌ Error al procesar los archivos: {ex}\n"
                print(Exception)

        loader.visible = False 
        page.update()
    

####################################### SECCIÓN FORMULARIO ####################################### 
    
    ######### ELEMENTOS DEL FORMULARIO ###############

    # Select de Marca
    select_marca = ft.Dropdown(
        enable_filter=True,
        editable=True,
        width=250,
        dense= 10,
        label="Marca",
        options= [ft.dropdown.Option(marca) for marca in sorted(mvnos)]
    )

    # Select sales
    select_pago_sales = ft.Dropdown(
        label = "Recibe comisión adelantada por sales",
        width=250,
        options=[ft.dropdown.Option("SI"), ft.dropdown.Option("NO")]
    )

    # Select proceso
    select_proceso = ft.Dropdown(
        label = "Proceso",
        width=250,
        options=[ft.dropdown.Option("Activación"), ft.dropdown.Option("Recarga")]
    )

    # Input comisión
    input_comision = ft.Dropdown(
        label = "Comisión %",
        width=250,
        options=[ft.dropdown.Option("20%"), ft.dropdown.Option("15%")]
    )

    input_fecha = ft.TextField(
        label="Introduce fecha dd-mm-aaaa",
        value="31-12-2025",
        width=250,
    )

    # Pickers para archivos CSV
    file_picker_1 = ft.FilePicker(on_result=on_file_picker_1_result)
    file_picker_2 = ft.FilePicker(on_result=on_file_picker_2_result)

    # Botón de procesamiento
    boton_procesar = ft.ElevatedButton(
        text="Procesar",
        icon=ft.Icons.PLAY_ARROW,
        on_click=procesar_archivos,
    )

    # Loader
    loader = ft.ProgressBar(visible=False)
    
    # Contenedor izquierdo (formulario)
    seccion_izquierda = ft.Column(
        [
            ft.Text("Carga de Archivos", weight="bold"),
            ft.Text("Selecciona el reporte de finanzas de la marca"),
            ft.ElevatedButton(
                "Seleccionar CSV 1", icon=ft.Icons.UPLOAD_FILE, on_click=lambda e: file_picker_1.pick_files()
            ),
            ft.Text("Selecciona el reporte de general ACT/REC"),
            ft.ElevatedButton(
                "Seleccionar Excel General", icon=ft.Icons.UPLOAD_FILE, on_click=lambda e: file_picker_2.pick_files()
            ),
            select_marca,
            select_pago_sales,
            select_proceso,
            input_comision,
            input_fecha,
            boton_procesar,
            loader,
        ],
        spacing=10,
        expand=1,
    )


###################################### SECCIÓN INFORMACIÓN #############################################
    
    ############## ELEMENTOS DE LA SECCION INFORMACIÓN ############
    
    # Controles de estado
    resumen_text = ft.Text(value="Carga tus archivos para ver información.", size=14)
    
    
    # Contenedor derecho (información)
    info_container = ft.Container(
    content=ft.Column(
            [
                ft.Text("Información", weight="bold"),
                ft.Container(resumen_text, bgcolor=ft.Colors.GREY_100),
            ],
            spacing=5,
            scroll=ft.ScrollMode.AUTO,  # Aquí debe ir
        ),
    height=600,
    width=800,
    padding=10,
    border_radius=10,
    bgcolor=ft.Colors.GREY_100,
)



############################################# LAYAOUT PRINCIPAL ##########################################
    # Layout principal
    page.title = "Gestor de Comisiones"
    page.window_width = 800
    page.window_height = 600
    page.scroll = "auto"
    page.window_icon = os.path.join("assets", "images", "saturn.ico")
    page.scroll = ft.ScrollMode.AUTO
    
    layout = ft.Row(
        [
            seccion_izquierda,
            info_container,
        ],
        expand=True,
    )

    page.overlay.append(file_picker_1)
    page.overlay.append(file_picker_2)
    page.add(layout)


ft.app(main)
