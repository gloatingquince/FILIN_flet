import math
from datetime import datetime, timedelta
import flet as ft
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import json

def calculate_sun_position(year, month, day, hour, minute, lat, lon, tz):
    """
    Высокоточный расчет положения Солнца (азимут и высота) без внешних библиотек.
    Погрешность около ±0.05° при условии корректного времени и координат.
    """
    # Преобразуем координаты
    lat_rad = math.radians(lat)
    lon = float(lon)

    # Дата и время
    dt = datetime(year, month, day, hour, minute)
    day_of_year = dt.timetuple().tm_yday
    time_utc = hour - tz + minute / 60.0

    # --- 1. Дробный год (в радианах) ---
    gamma = 2 * math.pi / 365 * (day_of_year - 1 + (time_utc - 12) / 24)

    # --- 2. Уравнение времени (в минутах) ---
    eqtime = 229.18 * (
        0.000075
        + 0.001868 * math.cos(gamma)
        - 0.032077 * math.sin(gamma)
        - 0.014615 * math.cos(2*gamma)
        - 0.040849 * math.sin(2*gamma)
    )

    # --- 3. Деклинация Солнца (в радианах) ---
    decl = (
        0.006918
        - 0.399912 * math.cos(gamma)
        + 0.070257 * math.sin(gamma)
        - 0.006758 * math.cos(2*gamma)
        + 0.000907 * math.sin(2*gamma)
        - 0.002697 * math.cos(3*gamma)
        + 0.00148 * math.sin(3*gamma)
    )

    # --- 4. Часовой угол (Hour Angle) ---
    time_offset = eqtime + 4 * lon
    tst = time_utc * 60 + time_offset
    ha = math.radians(tst / 4 - 180)

    # --- 5. Высота Солнца ---
    altitude_geom = math.degrees(
        math.asin(
            math.sin(lat_rad) * math.sin(decl)
            + math.cos(lat_rad) * math.cos(decl) * math.cos(ha)
        )
    )

    # --- 6. Атмосферная рефракция ---
    # Коррекция по формуле Bennett (точна до 15')
    if altitude_geom > -1:  # ниже -1° не применяем
        refraction = 1.02 / math.tan(math.radians(altitude_geom + 10.3/(altitude_geom + 5.11))) / 60
    else:
        refraction = 0
    altitude = altitude_geom + refraction

    # --- 7. Азимут ---
    azimuth = math.degrees(
        math.atan2(
            -math.sin(ha),
            math.tan(decl) * math.cos(lat_rad)
            - math.sin(lat_rad) * math.cos(ha)
        )
    )
    azimuth = (azimuth + 360) % 360

    return azimuth, altitude

def calculate_sun_trajectory(lat, lon, tz, start_dt, end_dt, step_minutes=1):
    result = []
    current_dt = start_dt
    while current_dt <= end_dt:
        az, alt = calculate_sun_position(
            current_dt.year, current_dt.month, current_dt.day,
            current_dt.hour, current_dt.minute,
            lat, lon, tz
        )
        result.append((current_dt, az, alt))
        current_dt += timedelta(minutes=step_minutes)
    return result

def main(page: ft.Page):
    page.title = "Солнечная траектория"
    page.bgcolor = "#EEEEEE"

    page.vertical_alignment = ft.MainAxisAlignment.START
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    page.window.min_width = 500
    page.window.max_width = 500
    page.window.width = 500

    page.window.min_height = 780
    page.window.max_height = 1000
    page.window.height = 1000

    page.window.resizable=True
    page.window.always_on_top=True

    #Создаем стили
    #ST0 Цвета
    st_color_BLACK=ft.Colors.BLACK
    st_color_WHITE=ft.Colors.WHITE
    st_color_GREY=ft.Colors="#EEEEEE"
    st_color_AcePink=ft.Colors="#F40059"
    st_color_PurePink=ft.Colors="#FFEBF3" 
    st_color_AcentPink=ft.Colors="#E6A8B5"

    #ST1 Стиль Кнопки
    BTN_AcePink=ft.ButtonStyle(
        bgcolor=st_color_PurePink,
        color=st_color_BLACK,
        shape=ft.RoundedRectangleBorder(radius=8),
        side=ft.BorderSide(width=1, color=st_color_AcePink)
    )
    
    BTN_DropDown=ft.ButtonStyle(
        bgcolor=st_color_WHITE,
        color=st_color_BLACK,
    )

    #ST0 Стили текста
    st_col_name=ft.TextStyle(
        color=st_color_BLACK,
        weight=ft.FontWeight.BOLD,
        size=14,
    )

    st_main_name=ft.TextStyle(
        color=st_color_BLACK,
        font_family="Inter",
        weight=ft.FontWeight.BOLD,
        size=18,
    )
    st_txt_entry=ft.TextStyle(
        color=st_color_BLACK,
        size=16,
    )

    st_second_name=ft.TextStyle(
        color=st_color_BLACK,
        size=20,
    )
    
    lat_field = ft.TextField(
        text_style=st_txt_entry,
        cursor_color=st_color_AcePink,
        focused_border_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=150,
        label="Широта",
        value="55.7558")
    
    lat_dir = ft.Dropdown(
        text_style=st_txt_entry,
        filled=True,
        fill_color=st_color_WHITE,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=150,
        label="Направление широты", 
        value="N",
        options=[
        ft.dropdown.Option("N",style=BTN_DropDown), ft.dropdown.Option("S",style=BTN_DropDown)
        ]
    )

    lon_field = ft.TextField(
        text_style=st_txt_entry,
        cursor_color=st_color_AcePink,
        focused_border_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=150,
        label="Долгота",
        value="37.6176")
    
    lon_dir = ft.Dropdown(
        text_style=st_txt_entry,
        filled=True,
        fill_color=st_color_WHITE,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=150,
        label="Направление долготы",
        value="E",
        options=[
        ft.dropdown.Option("E",style=BTN_DropDown), ft.dropdown.Option("W",style=BTN_DropDown)
        ]
    )

    tz_field = ft.TextField(text_style=st_txt_entry,
        cursor_color=st_color_AcePink,
        focused_border_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=100,
        label="Часовой пояс (UTC±)",
        value="3")

    start_date = ft.TextField(
        text_style=st_txt_entry,
        cursor_color=st_color_AcePink,
        focused_border_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=150,
        label="Дата начала (YYYY-MM-DD)",
        value="2025-08-04")
    
    start_time = ft.TextField(text_style=st_txt_entry,
        cursor_color=st_color_AcePink,
        focused_border_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=150,
        label="Время начала (HH:MM)",
        value="06:00")
    
    end_date = ft.TextField(text_style=st_txt_entry,
        cursor_color=st_color_AcePink,
        focused_border_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=150,
        label="Дата конца (YYYY-MM-DD)",
        value="2025-08-04")
    
    end_time = ft.TextField(text_style=st_txt_entry,
        cursor_color=st_color_AcePink,
        focused_border_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=150,
        label="Время конца (HH:MM)",
        value="20:00")
    
    step_field = ft.TextField(text_style=st_txt_entry,
        cursor_color=st_color_AcePink,
        focused_border_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        focused_border_width=1,
        bgcolor=st_color_WHITE,
        expand=True,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        width=100,
        label="Шаг (минуты)",
        value="60")

    result_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text(), heading_row_alignment=ft.MainAxisAlignment.START),
            ft.DataColumn(ft.Text(), heading_row_alignment=ft.MainAxisAlignment.START),
            ft.DataColumn(ft.Text(), heading_row_alignment=ft.MainAxisAlignment.START)
        ],
        rows=[],
        heading_row_height=10
    )


    def calculate(e):
        try:
            lat = float(lat_field.value)
            if lat_dir.value == "S":
                lat = -lat

            lon = float(lon_field.value)
            if lon_dir.value == "W":
                lon = -lon

            tz = float(tz_field.value)

            start_dt = datetime.strptime(start_date.value + " " + start_time.value, "%Y-%m-%d %H:%M")
            end_dt = datetime.strptime(end_date.value + " " + end_time.value, "%Y-%m-%d %H:%M")
            step = int(step_field.value)

            data = calculate_sun_trajectory(lat, lon, tz, start_dt, end_dt, step)
            # ⬇ Создаём DataFrame и сохраняем его в сессию
            df = pd.DataFrame(data, columns=['DateTime', 'Azimuth', 'Altitude'])
            page.session.set("last_result_df", df)

            result_table.rows.clear()
            for dt, az, alt in data:
                result_table.rows.append(
                    ft.DataRow(
                        cells=[
                            ft.DataCell(ft.Text(dt.strftime("%Y-%m-%d %H:%M"), selectable=True)),
                            ft.DataCell(ft.Text(f"{az:.3f}°", selectable=True)),
                            ft.DataCell(ft.Text(f"{alt:.3f}°", selectable=True)),
                        ]
                    )
                )
            page.update()

        except Exception as ex:
            page.snack_bar = ft.SnackBar(ft.Text(f"Ошибка: {ex}"), open=True)
            page.update()

    def handle_export_result(e: ft.FilePickerResultEvent):
        """Обрабатывает результат выбора файла для СОХРАНЕНИЯ"""
        if e.path:
            file_path = e.path
            export_format = page.session.get("export_format")

            df_sun_trajectory = page.session.get("last_result_df")
            
            try:
                if export_format == "json":
                    df_sun_trajectory.to_json(file_path, orient="records", indent=4, force_ascii=False)
                elif export_format == "csv":
                    df_sun_trajectory.to_csv(file_path, index=False, sep=";")
                elif export_format == "xml":
                    root = ET.Element("SunTrajectory")
                    for _, row in df_sun_trajectory.iterrows():
                        record = ET.SubElement(root, "PositionData")
                        dt_val = row["DateTime"]
                        # Если в ячейке хранится datetime, преобразуем в строку
                        if hasattr(dt_val, "strftime"):
                            dt_text = dt_val.strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            dt_text = str(dt_val)

                        az_text = str(row["Azimuth"])
                        alt_text = str(row["Altitude"])

                        ET.SubElement(record, "DateTime").text = dt_text
                        ET.SubElement(record, "Azimuth").text = az_text
                        ET.SubElement(record, "Altitude").text = alt_text

                    # Преобразуем в красиво отформатированную строку
                    rough_string = ET.tostring(root, 'utf-8')
                    reparsed = minidom.parseString(rough_string)
                    pretty_xml = reparsed.toprettyxml(indent="  ", encoding="utf-8")

                    # Записываем в файл
                    with open(file_path, "wb") as f:
                        f.write(pretty_xml)
            except Exception as ex:
                pass

    def handle_export_settings(e: ft.FilePickerResultEvent):
        if not e.path:
            return

        file_path = e.path
        settings = {
            "latitude": lat_field.value,
            "latitude_dir": lat_dir.value,
            "longitude": lon_field.value,
            "longitude_dir": lon_dir.value,
            "timezone": tz_field.value,
            "start_date": start_date.value,
            "start_time": start_time.value,
            "end_date": end_date.value,
            "end_time": end_time.value,
            "step": step_field.value,
        }

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
            page.snack_bar = ft.SnackBar(ft.Text(f"Настройки сохранены: {file_path}"), open=True)
        except Exception as ex:
            page.snack_bar = ft.SnackBar(ft.Text(f"Ошибка сохранения настроек: {ex}"), open=True)
        page.update()

    def handle_import_settings(e: ft.FilePickerResultEvent):
        if not e.files or not e.files[0].path:
            return

        file_path = e.files[0].path
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                settings = json.load(f)

            lat_field.value = settings.get("latitude", "")
            lat_dir.value = settings.get("latitude_dir", "N")
            lon_field.value = settings.get("longitude", "")
            lon_dir.value = settings.get("longitude_dir", "E")
            tz_field.value = settings.get("timezone", "")
            start_date.value = settings.get("start_date", "")
            start_time.value = settings.get("start_time", "")
            end_date.value = settings.get("end_date", "")
            end_time.value = settings.get("end_time", "")
            step_field.value = settings.get("step", "")

            page.snack_bar = ft.SnackBar(ft.Text(f"Настройки загружены: {file_path}"), open=True)
            page.update()
        except Exception as ex:
            page.snack_bar = ft.SnackBar(ft.Text(f"Ошибка импорта настроек: {ex}"), open=True)
            page.update()

    export_file_picker = ft.FilePicker(on_result=handle_export_result)
    settings_export_picker = ft.FilePicker(on_result=lambda e: handle_export_settings(e))
    settings_import_picker = ft.FilePicker(on_result=lambda e: handle_import_settings(e))
    page.overlay.extend([settings_export_picker, settings_import_picker])
    page.overlay.extend([export_file_picker])

    def export_data(e, format_type: str):
        """Вызывается по нажатию кнопки в диалоге экспорта"""
        page.session.set("export_format", format_type)
        page.update()
        
        export_file_picker.save_file(
            dialog_title=f"Сохранить как {format_type.upper()}",
            file_name=f"SunTrack.{format_type}",
            allowed_extensions=[format_type]
        )

    def open_export_dialog(e):
        """Открывает диалог выбора формата для экспорта"""
        dialog_exp = ft.AlertDialog(
            title=ft.Text("Выберите формат экспорта", style=st_second_name),
            bgcolor=st_color_GREY,
            content=ft.Column(
                [
                    ft.OutlinedButton("Экспорт в JSON (.json)", style=BTN_AcePink,on_click=lambda _: export_data(_, "json"), width=250),
                    ft.OutlinedButton("Экспорт в CSV (.csv)", style=BTN_AcePink,on_click=lambda _: export_data(_, "csv"), width=250),
                    ft.OutlinedButton("Экспорт в XML (.xml)", style=BTN_AcePink,on_click=lambda _: export_data(_, "xml"), width=250),
                ],
                tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER
            )
        )
        dialog_exp.open=True
        page.overlay.append(dialog_exp)
        page.update()

    def export_settings(e):
        """Вызывает диалог сохранения JSON с настройками"""
        settings_export_picker.save_file(
            dialog_title="Сохранить настройки",
            file_name="sun_settings.json",
            allowed_extensions=["json"]
        )


    def import_settings(e):
        """Вызывает диалог выбора JSON с настройками"""
        settings_import_picker.pick_files(
            dialog_title="Выбрать файл настроек",
            allowed_extensions=["json"],
            allow_multiple=False
        )


    page.add(
        ft.Text("Расчет траектории солнца", style=st_main_name),
        ft.Container(
            bgcolor=st_color_WHITE, 
            width = 490,
            height=900,
            padding=ft.padding.only(top=10,left=5,right=5,bottom=5),
            expand=True, 
            border_radius=ft.border_radius.all(8),
            border=ft.border.all(2, color=st_color_AcePink),
            content=ft.Column([
                ft.Row([
                    ft.OutlinedButton("Экспорт настроек", style=BTN_AcePink, on_click=export_settings),
                    ft.OutlinedButton("Импорт настроек", style=BTN_AcePink,on_click=import_settings),
                    ],height=25, spacing=20, vertical_alignment=ft.CrossAxisAlignment.CENTER, alignment=ft.MainAxisAlignment.START),
                ft.Divider(height=10, leading_indent=0, thickness=1, color=st_color_AcePink),
                ft.Row([
                    ft.Column([lat_field,lon_field,start_date,end_date],spacing=20, alignment=ft.MainAxisAlignment.START,  horizontal_alignment=ft.CrossAxisAlignment.CENTER,),
                    ft.Column([lat_dir,lon_dir,start_time,end_time],spacing=20, alignment=ft.MainAxisAlignment.START, horizontal_alignment=ft.CrossAxisAlignment.CENTER,),
                    ft.Column([tz_field,step_field], alignment=ft.MainAxisAlignment.SPACE_AROUND, horizontal_alignment=ft.CrossAxisAlignment.CENTER,),
                ],height=250, spacing=20, vertical_alignment=ft.CrossAxisAlignment.CENTER, alignment=ft.MainAxisAlignment.CENTER),
                ft.Divider(height=2, leading_indent=0, thickness=1, color=st_color_AcePink),
                ft.Row([
                    ft.OutlinedButton("Расчет", style=BTN_AcePink, on_click=calculate),
                    ft.OutlinedButton("Экспорт тракетории", style=BTN_AcePink, on_click=open_export_dialog),
                ],height=25, spacing=20,vertical_alignment=ft.CrossAxisAlignment.CENTER, alignment=ft.MainAxisAlignment.START),
                ft.Divider(height=2, leading_indent=0, thickness=1, color=st_color_AcePink),
                ft.Row([ft.Text("Дата/время", style=st_col_name),
                        ft.Text("Азимут",style=st_col_name),
                        ft.Text("Высота",style=st_col_name)
                        ],height=30,vertical_alignment=ft.CrossAxisAlignment.START, alignment=ft.MainAxisAlignment.SPACE_EVENLY),
                ft.Column([result_table],alignment=ft.MainAxisAlignment.START, horizontal_alignment=ft.CrossAxisAlignment.CENTER, height=600,scroll=ft.ScrollMode.HIDDEN)
            ], width=480, horizontal_alignment=ft.CrossAxisAlignment.CENTER, alignment=ft.MainAxisAlignment.START),
        )
    )

if __name__ == "__main__":
    ft.app(target=main)


