import flet as ft
import os
import json
import subprocess

# Global lists for storing data
NWC_CONVERSION_MOVE_COMMAND=""
NWC_MOVE_COMMAND = []

def main(page: ft.Page):
    page.bgcolor = "#EEEEEE"
    page.title = "FILIN BAT Master"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window.min_width = 850
    page.window.max_width = 860
    page.window.width = 850
    page.window.min_height = 900
    page.window.height = 900
    page.auto_scroll=True,

    model_names = []
    path_to_ref_each = []
    SET_NAMES = []
    Name_TXT_SET = []

    Multi=False
    
    #Создаем стили
    #ST0 Цвета
    st_color_BLACK=ft.Colors.BLACK
    st_color_BerryPink=ft.Colors="#D0034E"
    st_color_AcePink=ft.Colors="#F40059"
    st_color_PurePink=ft.Colors="#FFEBF3"
    st_color_BgGrey="#EEEEEE"
    st_color_label="#57636c"
    st_color_txtGrey="#060606"

    #UIСоздать стили UI
    txt_entry_names=ft.TextStyle(
        size=18,
        weight=ft.FontWeight.W_500,
        color=st_color_BLACK,
    )
    #UIСоздать Элементы UI
    AcePink_BTN=ft.ButtonStyle(
        bgcolor=st_color_PurePink,
        color=st_color_txtGrey,
        shape=ft.RoundedRectangleBorder(radius=8),
        side=ft.BorderSide(width=1, color=st_color_AcePink)
    )

    #UI1 Блок RVT Path
    #UI1 Создать FilePicker RVT Path
    def path_to_ref(e: ft.FilePickerResultEvent):
        nonlocal Multi
        selected_ref_path.value = (
            (e.path) if e.path else "Пути не получены(("
            )
        selected_ref_path.update()
        Multi=False
        get_path_bat_name()
 
    #UI1 Создать Событие FilePicker RVT Path
    get_ref_path = ft.FilePicker(on_result=path_to_ref)
    selected_ref_path = ft.TextField(
        focused_border_color=st_color_AcePink,
        cursor_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        bgcolor="white",
    )
    page.overlay.append(get_ref_path)
    
    #UI1 Создать Кнопку Выбора RVT Path
    btn_get_ref_path=ft.OutlinedButton(
        text="Выбор папки",
        height=30,
        style=AcePink_BTN,
        on_click=lambda _:get_ref_path.get_directory_path(
            dialog_title= "Путь до REF Файлов"
            )
    )
    
    #UI1 Создать FilePicker multi RVT Path
    def file_picker_multi_ref_on_result(e: ft.FilePickerResultEvent):
        nonlocal Multi, path_to_ref_each, model_names
        
        if e.files:
            Multi=True
            model_names.clear()
            path_to_ref_each.clear()
            model_names = [f.name for f in e.files]
            path_to_ref_each = [f.path for f in e.files]

            log_text.value = f"Выбраны файлы:\n {model_names}\n{path_to_ref_each}"
            page.update()
        else:
            selected_ref_path.value = "Файлы не выбраны(("
            Multi=False
            page.update()

    #UI1 Создать Событие FilePicker multi RVT Path
    file_picker_multi_ref = ft.FilePicker(on_result=file_picker_multi_ref_on_result)
    page.overlay.append(file_picker_multi_ref)

    #UI1 Создать Кнопку Выбора multi RVT 
    btn_MultiFile_Ref = ft.OutlinedButton(
        "Выбор файла",
        height=30,
        style=AcePink_BTN,
        on_click=lambda e: file_picker_multi_ref.pick_files(allow_multiple=True)
    )


    #L1 Получить именя файлов REF из пути selected_ref_path
    def get_path_bat_name():
        if Multi is False:
            if selected_ref_path:
                model_names.clear()
                path_to_ref_each.clear()
                try:
                    files = os.listdir(selected_ref_path.value)
                    for file in files:
                        if os.path.isfile(os.path.join(selected_ref_path.value, file)):
                            name, ext = os.path.splitext(file)
                            model_names.append(name)
                            path_to_ref_each.append(os.path.join(selected_ref_path.value, file))
                except Exception as Ex:
                    pass
            print(model_names)
            generate_set_names()
        
    #L2 Генерация содержимого BAT
    #L2.1 Генерация переменных имен "set name{i}=..."
    def generate_set_names():
    
        SET_NAMES.clear()
        for index, name in enumerate(model_names):
            SET_NAMES.append(f"set name{index}={name}")
        print(SET_NAMES)
    #UI2 Блок NWC Path
    #UI2 Создать FilePicker NWC Path
    def path_to_nwc(e: ft.FilePickerResultEvent):
        selected_nwc_path.value = (e.path) if e.path else "Пути не получены(("
        selected_nwc_path.update()
    #UI2 Создать Событие FilePicker NWC Path
    get_nwc_path = ft.FilePicker(on_result=path_to_nwc)
    selected_nwc_path = ft.TextField(
        focused_border_color=st_color_AcePink,
        cursor_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        bgcolor="white",
    )
    page.overlay.append(get_nwc_path)

    #UI2 Создать Кнопку Выбора NWC Path
    btn_get_nwc_path=ft.OutlinedButton(
        text="Выбрать",
        height=30,
        style=ft.ButtonStyle(
            bgcolor="#FFEBF3",
            color="#060606",
            shape=ft.RoundedRectangleBorder(radius=8),
            side=ft.BorderSide(width=1, color=st_color_AcePink)
        ),
        on_click=lambda _: get_nwc_path.get_directory_path(
            dialog_title= "Путь до NWC Файлов"
            )
    )
    
    #UI3 Блок BAT
    #UI3 Создать FilePicker BAT Path
    def path_to_BAT(e: ft.FilePickerResultEvent):
        selected_BAT_path.value = (e.path) if e.path else "Пути не получены(("
        selected_BAT_path.update()
    #UI3 Создать Событие FilePicker BAT Path
    get_BAT_path = ft.FilePicker(on_result=path_to_BAT)
    selected_BAT_path = ft.TextField(
        focused_border_color=st_color_AcePink,
        cursor_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        bgcolor="white",
    )
    page.overlay.append(get_BAT_path)

    #UI3 Создать Кнопку выбора пути сохранения BAT
    btn_get_BAT_path=ft.OutlinedButton(
        text="Выбрать",
        height=30,
        style=ft.ButtonStyle(
            bgcolor="#FFEBF3",
            color="#060606",
            shape=ft.RoundedRectangleBorder(radius=8),
            side=ft.BorderSide(width=1, color=st_color_AcePink)
        ),
        on_click=lambda _: get_BAT_path.get_directory_path(
            dialog_title= "Путь cохранения BAT"
            )
    )
    
    #UI3 Создать поле ввода версии NW
    NW_version=ft.TextField(
        focused_border_color=st_color_AcePink,
        cursor_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        bgcolor="white"
    )
    
    #UI3 Создать поле ввода версии Экспортера NW
    NW_exporter_version=ft.TextField(
        focused_border_color=st_color_AcePink,
        cursor_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        bgcolor="white"
    )

    #UI3 Создать поле ввода имени BAT
    bat_Name_enrty=ft.TextField(
        focused_border_color=st_color_AcePink,
        cursor_color=st_color_AcePink,
        value="bat_file.bat",
        filled=True,
        border_radius=8,
        bgcolor="white",
    )

    #UI4 Создать Элементы блока сохранть BAT.
    #L2.1 Генерация переменной имени конфига для экспортера "set name{i}=..."
    def generate_name_txt_set():
        Name_TXT_SET.clear()
        Name_TXT_SET.append(f"set name_txt={bat_Name_enrty.value}")
        print(Name_TXT_SET)
    #L2.2 Генерация TXT конфига для NWS
    def export_to_txt():
        filename = (bat_Name_enrty.value) + ".txt"
        filepath = os.path.join(selected_BAT_path.value, filename)
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                for path in path_to_ref_each:
                    f.write(path + "\n")
        except Exception as Ex:
            pass
        print(f"Name_TXT_SET={filepath}")
    
    #L2.3 Генерация комарнды экспорта NWC
    def create_NWC_CONVERSION_MOVE_COMMAND():
        global NWC_CONVERSION_MOVE_COMMAND
        NWC_CONVERSION_MOVE_COMMAND=""
        command = f'"C:\\Program Files\\Autodesk\\Navisworks Manage {NW_version.value}\\FiletoolsTaskRunner.exe" /i "{selected_BAT_path.value}\\%name_txt%.txt" /of "C:\\Temp\\%name_txt%.nwd" /version {NW_exporter_version.value}'
        NWC_CONVERSION_MOVE_COMMAND=(command)
        print(NWC_CONVERSION_MOVE_COMMAND)
    
    #L2.4 Генерация комарнды переноса NWC
    def generate_nwc_move_command():
        global NWC_MOVE_COMMAND
        NWC_MOVE_COMMAND.clear()
        for index, name in enumerate(model_names):
            command = f'move "{selected_ref_path.value}\\%name{index}%.nwc" "{selected_nwc_path.value}\\%name{index}%.nwc"'  # Corrected this line
            NWC_MOVE_COMMAND.append(command)
        print(NWC_MOVE_COMMAND)

    # Генерация Bat файла
    #L0 неизменные переменные содержимого BAT
    Start_bat = "chcp 65001>nul" #Условипе запуска BAT
    Temp_Buffer = "MD C:\\temp" #Папка временного хранения
    TIMEOUT_BAT = "timeout 300" #Время до закрытия консоли полсе отработки
    #UI4 Создать определение охранть BAT

    def gen_bat(e):
        """
        Creates a .bat file with the specified name, path, encoding, and line ending.
        The file content is structured as follows:
        1. Start_bat
        2. SET_NAMES (each item on a new line)
        3. Name_TXT_SET (if not empty)
        4. Temp_Buffer
        5. NWC_CONVERSION_MOVE_COMMAND (if Navisworks and exporter versions are provided)
        6. NWC_MOVE_COMMAND (each command on a new line)
        7. TIMEOUT_BAT
        Gets all the values from the Entry fields.
        """
        encoding = "utf-8"
        line_ending = "\r\n"  # Set line ending to Windows (CRLF)

        generate_set_names()
        generate_name_txt_set()
        export_to_txt()
        create_NWC_CONVERSION_MOVE_COMMAND()
        generate_nwc_move_command()

        filepath = os.path.join(selected_BAT_path.value, bat_Name_enrty.value + ".bat")
        try:
            with open(filepath, "w", encoding=encoding, newline=line_ending) as f:
                f.write(f"{Start_bat}\n")
                for set_name in SET_NAMES:
                    f.write(f"{set_name}\n")
                if Name_TXT_SET:
                    f.write(f"{Name_TXT_SET[0]}\n")
                f.write(f"{Temp_Buffer}\n")
                if (NW_version.value) and (NW_exporter_version.value):
                    f.write(f"{NWC_CONVERSION_MOVE_COMMAND}\n")
                for move_command in NWC_MOVE_COMMAND:
                    f.write(f"{move_command}\n")
                f.write(f"{TIMEOUT_BAT}\n")
            log_text.value=f'Содержание Bat файла: \n {Start_bat}\n {SET_NAMES}\n {Name_TXT_SET[0]}\n {Temp_Buffer}\n {NWC_CONVERSION_MOVE_COMMAND}\n {NWC_MOVE_COMMAND}\n {TIMEOUT_BAT}\n'
            log_text.update()
        except Exception as EX:
            pass
        
    #UI4 Создать Кнопку сохранть H.
    save_bat=ft.OutlinedButton(
        text="Сохранить BAT",
        width=850,
        height=40,
        style=ft.ButtonStyle(
            bgcolor="#FFEBF3",
            color=st_color_BerryPink,
            shape=ft.RoundedRectangleBorder(radius=8),
            side=ft.BorderSide(width=1, color=st_color_BerryPink)
        ),
        on_click=gen_bat
    )

    #UI4.1 Создать FilePickerResault Сохранть Настройки
    def save_settings_res(e: ft.FilePickerResultEvent):
        save_path=e.path
        save_name=e.name
        filepath = os.path.join(save_path + "_MasterBat.json")
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        settings = {
            "navisworks_version": NW_version.value,
            "exporter_version": NW_exporter_version.value,
        }
        if filepath:
            with open(filepath,"w", encoding="utf-8") as f:
                        json.dump(settings, f, indent=4)
        page.update()
    #UI4.1 Создать FilePicker Сохранть Настройки
    save_settings = ft.FilePicker(on_result=save_settings_res)
    page.overlay.append(save_settings)
    page.update()

    #UI4.1 Создать Кнопку Сохранть Настройки.
    btn_save_setting=ft.OutlinedButton(
        text="Сохранить настройки",
        width=190,
        height=30,
        style=AcePink_BTN,
        on_click=lambda _: save_settings.save_file(
            dialog_title= "Сохранение настроек",
            allowed_extensions=["json"]
            )
    )

    #UI4.2 Создать FilePickerResault Загрузить Настройки
    def load_settings_result(e: ft.FilePickerResultEvent):
        for f in load_settings.result.files:
            name=f.name
            path=f.path
        with open(path, "r") as f:
            settings = json.load(f)
            NW_version.value=(settings['navisworks_version'])
            NW_exporter_version.value=(settings['exporter_version'])
            NW_version.update()
            NW_exporter_version.update()
    #UI4.2 Создать FilePicker Загрузить Настройки
    load_settings = ft.FilePicker(on_result=load_settings_result)
    page.overlay.append(load_settings)
    page.update()

    #UI4.2 Создать Кнопку Загрузить Настройки.
    btn_load_setting=ft.OutlinedButton(
        text="Загрузить настройки",
        width=190,
        height=30,
        style=AcePink_BTN,
        on_click=lambda _: load_settings.pick_files(
            dialog_title= "Загрузить настройки",
            allowed_extensions=["json"],
            allow_multiple=False
            )
    )

    log_text=ft.TextField(
        focused_border_color=st_color_AcePink,
        cursor_color=st_color_AcePink,
        multiline=True,
        min_lines=5,
        read_only=True,
        border_radius=8,
        bgcolor="white",
        value=f'Содержание Bat файла появиться при создании.'                      
    )
    
    #UI1 Создать Событие  "показать бат"
    def show_bat_folder(e):
        path = selected_BAT_path.value.strip()
        print(path)
        subprocess.run(f'explorer "{os.path.abspath(path)}"', shell=True)
        page.update()
        
    #UI1 Создать Кнопку "показать бат"
    btn_open_bat_folder = ft.OutlinedButton(
        "Показать bat",
        width=190,
        height=30,
        style=AcePink_BTN,
        on_click=show_bat_folder
    )


    #UI Кнопка Bat Planner
    #UI Событие
    def BTN_PlannerBat_Click(e):
        FPB="FILIN_BAT_Planner.exe"
        one_dir = os.path.dirname(os.path.abspath(__file__))
        full_path = os.path.join(one_dir, FPB)

        try:
            if not os.path.exists(full_path):
                full_path = FPB

            result = subprocess.run(
                    [full_path],
                    capture_output=True,
                    text=True
                )

            if result.returncode != 0:
                print(f"Произошла ошибка при выполнении {FPB}. Код возврата: {result.returncode}")
                print(f"stderr: {result.stderr}")
            else:
                print(result.stdout)

        except Exception as Ex:
            print(f"Исключение при запуске {FPB}: {Ex}")
    
    #UI Кнопка
    BTN_Planner_bat=ft.OutlinedButton(
        text="BAT Planner",
        width=190,
        height=30,
        style=AcePink_BTN,
        on_click = BTN_PlannerBat_Click,
    )

    #UI4.2 Создать Text Field для инструкции.
    Instruction_TF=ft.TextField(
        multiline=True,
        min_lines=5,
        read_only=True,
        filled=True,
        border_radius=8,
        bgcolor="white",
        text_size=16,
        color=st_color_BLACK,
        value=f'Утилита предназначена для генерации .bat файла, осуществляющего фоновый экспорт NWC.\n\nПо умолчанию Bat создается с разметкой Windows(CRLF) и кодировкой UTF-8.\n\n"Сохранить настройки" позволяет сохранить, предварительно заполненные, Версию Navisworks и Версию Экспортера Navisworks в .json, "Загрузить настройки" соответвенно загружает сохраненные настройки." '
    )

    form_MasterBAT=ft.Tabs(
        width=900,
        selected_index=0,
        animation_duration=300,
        indicator_color=st_color_BerryPink,
        label_color=st_color_BerryPink,
        unselected_label_color=st_color_label,
        adaptive=True,
        scrollable=True,
        label_text_style = ft.TextStyle(size=24, weight=ft.FontWeight.W_800),
        tabs=[
            ft.Tab(
                text="BAT Master",
                content=ft.Container(
                    padding=ft.padding.only(top=10),
                    width=700,
                    height=900,
                    alignment=ft.alignment.top_center,
                    bgcolor=st_color_BgGrey,
                    border_radius=ft.border_radius.all(20),
                    content=ft.Column(
                        scroll=ft.ScrollMode.AUTO,
                        expand=True,
                        alignment=ft.MainAxisAlignment.SPACE_EVENLY,
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Container(
                                        width=400,
                                        height=100,
                                        bgcolor=st_color_BgGrey,
                                        border_radius=ft.border_radius.all(5),
                                        content=ft.Column(
                                            alignment=ft.MainAxisAlignment.SPACE_EVENLY,
                                            controls=[
                                                ft.Row(
                                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                                    controls=[
                                                        ft.Text("Путь к файлам RVT", 
                                                            style=txt_entry_names,
                                                        ),
                                                        btn_get_ref_path,
                                                        btn_MultiFile_Ref
                                                    ]
                                                ),
                                                selected_ref_path
                                            ]
                                        )
                                    ),
                                    ft.Container(
                                        width=400,
                                        height=100,
                                        bgcolor=st_color_BgGrey,
                                        border_radius=ft.border_radius.all(5),
                                        content=ft.Column(
                                            alignment=ft.MainAxisAlignment.SPACE_EVENLY,
                                            controls=[
                                                ft.Text("Версия Navisworks", 
                                                    style=txt_entry_names,
                                                ),
                                                NW_version
                                            ]
                                        )
                                    )
                                ]
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Container(
                                        width=400,
                                        height=100,
                                        bgcolor=st_color_BgGrey,
                                        border_radius=ft.border_radius.all(5),
                                        content=ft.Column(
                                            alignment=ft.MainAxisAlignment.SPACE_EVENLY,
                                            controls=[
                                                ft.Row(
                                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                                    controls=[
                                                        ft.Text("Путь к файлам NWC",
                                                            style=txt_entry_names
                                                        ),
                                                        btn_get_nwc_path
                                                    ]
                                                ),
                                                selected_nwc_path
                                            ]
                                        )
                                    ),
                                    ft.Container(
                                        width=400,
                                        height=100,
                                        bgcolor=st_color_BgGrey,
                                        border_radius=ft.border_radius.all(5),
                                        content=ft.Column(
                                            alignment=ft.MainAxisAlignment.SPACE_EVENLY,
                                            controls=[
                                                ft.Text("Версия экспортера Navisworks",
                                                    style=txt_entry_names
                                                ),
                                                NW_exporter_version
                                            ]
                                        )
                                    )
                                ]
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Container(
                                        width=400,
                                        height=100,
                                        bgcolor=st_color_BgGrey,
                                        border_radius=ft.border_radius.all(5),
                                        content=ft.Column(
                                            alignment=ft.MainAxisAlignment.SPACE_EVENLY,
                                            controls=[
                                                ft.Row(
                                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                                    controls=[
                                                        ft.Text("Путь сохранения BAT",
                                                            style=txt_entry_names
                                                        ),
                                                        btn_get_BAT_path
                                                    ]
                                                ),
                                                selected_BAT_path
                                            ]
                                        )
                                    ),
                                    ft.Container(
                                        width=400,
                                        height=100,
                                        bgcolor=st_color_BgGrey,
                                        border_radius=ft.border_radius.all(5),
                                        content=ft.Column(
                                            alignment=ft.MainAxisAlignment.SPACE_EVENLY,
                                            controls=[
                                                ft.Text("Имя BAT",
                                                    style=txt_entry_names
                                                ),
                                                bat_Name_enrty
                                            ]
                                        )
                                    )
                                ]
                            ),
                            ft.Container(
                                alignment=ft.alignment.center,
                                width=850,
                                height=40,
                                content=save_bat
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    btn_save_setting,
                                    btn_load_setting,
                                    btn_open_bat_folder,
                                    BTN_Planner_bat
                                ]
                            ),
                            ft.Container(
                                alignment=ft.alignment.center_left,
                                width=850,
                                height=40,
                                content=ft.Text("ЛОГ:", size=20, weight=ft.FontWeight.W_600, text_align=ft.TextAlign.LEFT,color=st_color_BLACK),
                            ),                            
                            log_text
                        ]
                    )
                ),
            ),
            ft.Tab(
                text="Инструкция",
                content=ft.Container(
                    padding=ft.padding.only(top=10),
                    width=700,
                    height=900,
                    alignment=ft.alignment.top_center,
                    bgcolor=st_color_BgGrey,
                    border_radius=ft.border_radius.all(20),
                    content=Instruction_TF,
                )
            ),
        ],
    expand=True,
    )

    #Добавить Элементы UI на страницу
    page.add(
        ft.Column(
            spacing=5,
            alignment = ft.MainAxisAlignment.CENTER,
            controls=[
                ft.Container(
                    content=form_MasterBAT,
                    alignment=ft.alignment.top_center,
                    bgcolor=st_color_BgGrey,
                    width=900,
                    height=900,
                    border_radius=10,
                )
            ]    
        )
    ) 

    #Создать DEF ON_click

    

ft.app(main)