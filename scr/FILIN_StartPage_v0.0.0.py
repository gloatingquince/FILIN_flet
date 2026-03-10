import flet as ft
import subprocess
import sys
import os
import json

def main(page: ft.Page):
    page.title = "FILIN"
    page.bgcolor="#FFCDE3"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window.min_width = 250
    page.window.max_width = 250
    page.window.width = 250
    page.window.min_height = 550
    page.window.height = 550
    page.auto_scroll=True
    page.window.always_on_top=True
    page.window.icon=ft.Image(src="images/FILIN_ICO.png", color=ft.Colors.BLACK, width=70, height=70)


    #Создаем стили
    #ST0 Цвета
    st_color_BLACK=ft.Colors.BLACK
    st_color_WHITE=ft.Colors.WHITE
    st_color_BerryPink=ft.Colors="#D0034E"
    st_color_PurePink=ft.Colors="#FFEBF3"
    st_color_Pink=ft.Colors="#FFD6E8"
    st_color_CandyFloss=ft.Colors="#F0F9F8"
    st_color_DarkMint=ft.Colors="#00685D"

    #ST1 Стиль Кнопки
    st_BTN_MainMenu=ft.ButtonStyle(
        bgcolor=st_color_PurePink,
        color=st_color_BLACK,
        shape=ft.RoundedRectangleBorder(radius=8),
        side=ft.BorderSide(width=2, color=st_color_BerryPink),
    )

    st_BTN_Urls=ft.ButtonStyle(
        bgcolor=st_color_CandyFloss,
        color=st_color_DarkMint,
        shape=ft.RoundedRectangleBorder(radius=8),
        side=ft.BorderSide(width=2, color=st_color_DarkMint),
        text_style=ft.TextStyle(weight=ft.FontWeight.W_500, decoration=ft.TextDecoration.UNDERLINE)
    )
    
    st_BTN_AllertDi=ft.ButtonStyle(
        color=st_color_BerryPink,
    )

    #ST1 Стиль текст
    st_txt_LOGO=ft.TextStyle(
        size=30,
        color=st_color_BLACK,
        weight=ft.FontWeight.W_900,
        font_family="Inter"
    )

    st_txt_txtfield=ft.TextStyle(
        size=12,
        color=st_color_BLACK,
    )

    #fn НАСТРОЙКИ
    settings_file = "settings.json"
    
    eir_launch_url = ('')
    bep_launch_url = ('')
    
    eir_url_field = ft.TextField(value=eir_launch_url, text_style=st_txt_txtfield, text_vertical_align=0, width=300, height=35,content_padding=1, cursor_height=15, focused_border_color=st_color_BerryPink , cursor_color=st_color_BerryPink)
    bep_url_field = ft.TextField(value=bep_launch_url, text_style=st_txt_txtfield, text_vertical_align=0, width=300, height=35,content_padding=1, cursor_height=15, focused_border_color=st_color_BerryPink , cursor_color=st_color_BerryPink)

    def save_settings(_=None):
        nonlocal eir_launch_url, bep_launch_url
        eir_launch_url = eir_url_field.value
        bep_launch_url = bep_url_field.value
        data = {
            "eir_launch_url": eir_launch_url,
            "bep_launch_url": bep_launch_url
        }
        with open(settings_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        settings_dialog.open = False
        page.update()

    def load_settings():
        nonlocal eir_launch_url, bep_launch_url
        try:
            with open(settings_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                eir_launch_url = data.get("eir_launch_url", "")
                bep_launch_url = data.get("bep_launch_url", "")

                eir_url_field.value = eir_launch_url
                bep_url_field.value = bep_launch_url
                page.update()
        except FileNotFoundError:
            pass

    def open_settings_dialog(e=None):
        load_settings()
        settings_dialog.open = True
        page.open(settings_dialog)
        page.update()

            
    #UI4.2 Создать FilePicker Загрузить Настройки

    settings_dialog = ft.AlertDialog(
        open=False,
        bgcolor=st_color_WHITE,
        actions_alignment=ft.MainAxisAlignment.END,
        actions_padding = 5,
        content_padding = 5,
        title_padding = 5,
        shape=ft.RoundedRectangleBorder(radius=10),
        modal=True,
        title=ft.Text("Настройки",
            size=16,
            weight=ft.FontWeight.W_700,
            color=st_color_BLACK,
            font_family="Inter"
        ),

        content=ft.Column(
            controls=[
                eir_url_field,
                bep_url_field
            ],
            tight=True
        ),

        actions=[
            ft.TextButton("Сохранить", height=20, style=st_BTN_AllertDi, on_click=save_settings),
            ft.TextButton("Отмена", height=20, style=st_BTN_AllertDi, on_click=lambda e: close_dialog()),
        ],
        
    )

    def close_dialog():
        settings_dialog.open = False
        page.update()

    settings_BTN=ft.TextButton(
        text="Настройки",
        width=100,
        height=20,
        style=st_BTN_AllertDi,
        on_click = open_settings_dialog,
    )

    #UI Блок ЛОГО
    logo_block = ft.Container(
        width=250,
        height=80,
        alignment=ft.alignment.center,
        border_radius=ft.border_radius.all(0),
        border=ft.border.all(1, color=st_color_WHITE),
        padding=0,
        content=ft.Row(
            controls=[
                ft.Image(src="images/FILIN_ICO.png", width=70, height=70),
                ft.Column(
                    controls=[
                        ft.Text(
                            "FILIN",
                            style=st_txt_LOGO,
                        ),
                        settings_BTN
                    ],
                    spacing=1,
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER
                )
            ],
            spacing=5,
            alignment=ft.MainAxisAlignment.CENTER,
            vertical_alignment=ft.CrossAxisAlignment.CENTER
        ),
    )
    
    #UI0 Кнопки
    #UI1 Кнопки Функции

    #UI1.1 Кнопка Master Bat
    #UI Событие 
    def BTN_MasterBat_Click(e):
        FMB = "FILIN_BATMaster.exe"
        one_dir = os.path.dirname(os.path.abspath(__file__))
        full_path = os.path.join(one_dir, FMB)

        try:
            if not os.path.exists(full_path):
                full_path = FMB

            result = subprocess.run(
                    [full_path],
                    capture_output=True,
                    text=True
                )

            if result.returncode != 0:
                print(f"Произошла ошибка при выполнении {FMB}. Код возврата: {result.returncode}")
                print(f"stderr: {result.stderr}")
            else:
                print(result.stdout)

        except Exception as Ex:
            print(f"Исключение при запуске {FMB}: {Ex}")

    #UI Кнопка        
    BTN_Master_bat=ft.OutlinedButton(
        text="BAT Master",
        width=200,
        height=40,
        style=st_BTN_MainMenu,
        on_click = BTN_MasterBat_Click,
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
        width=200,
        height=40,
        style=st_BTN_MainMenu,
        on_click = BTN_PlannerBat_Click,
    )

    #UI Кнопка Hotkey rvt
    #UI Событие
    def BTN_Hotkey_rvt_Click(e):
        FPB="FILIN_SHORTCUT_RVT.exe"
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
    BTN_Hotkey_rvt=ft.OutlinedButton(
        text="Hotkeys RVT",
        width=200,
        height=40,
        style=st_BTN_MainMenu,
        on_click = BTN_Hotkey_rvt_Click,
    )

    #UI Кнопка Matrx
    #UI Событие
    def BTN_Matrix_Click(e):
        FPB="FILIN_Matrix.exe"
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
    BTN_Matrix=ft.OutlinedButton(
        text="Matrix",
        width=200,
        height=40,
        style=st_BTN_MainMenu,
        on_click = BTN_Matrix_Click,
    )

    #UI Кнопка SunPath
    #UI Событие
    def BTN_SunPath_Click(e):
        FPB="FILIN_Matrix.exe"
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
    BTN_SunPath=ft.OutlinedButton(
        text="SunPath",
        width=200,
        height=40,
        style=st_BTN_MainMenu,
        on_click = BTN_SunPath_Click,
    )



    #UI Кнопка Open EIR
    #UI Событие
    def BTN_EIR_Click(e):
        load_settings()
        url=eir_launch_url
        if url:
            page.launch_url(url)


    #UI Кнопка
    BTN_Open_EIR=ft.OutlinedButton(
        text="EIR",
        width=50,
        height=40,
        style=st_BTN_Urls,
        on_click = BTN_EIR_Click,
    )

    #UI Кнопка Open BEP
    #UI Событие
    def BTN_BEP_Click(e):
        load_settings()
        url=bep_launch_url
        if url:
            page.launch_url(url)

    #UI Кнопка
    BTN_Open_BEP=ft.OutlinedButton(
        text="BEP",
        width=50,
        height=40,
        style=st_BTN_Urls,
        on_click = BTN_BEP_Click,
    )

    #UI Кнопка Open BEP
    #UI Событие
    def BTN_Open_GIT_Click(e):
        page.launch_url('https://github.com/gloatingquince')
        

    #UI Кнопка
    BTN_Open_GIT=ft.OutlinedButton(
        text="github",
        height=40,
        style=st_BTN_Urls,
        on_click = BTN_Open_GIT_Click,
    )

    #UI2 Группы конролов.

    URL_row=ft.Row([BTN_Open_EIR,
            BTN_Open_BEP,
            BTN_Open_GIT
        ], width=190, alignment=ft.MainAxisAlignment.SPACE_BETWEEN, vertical_alignment=ft.CrossAxisAlignment.CENTER,
    )

    column_buttons = ft.Container(
        width=250,
        bgcolor=st_color_WHITE,
        alignment=ft.alignment.bottom_center,
        border_radius=ft.border_radius.all(10),
        border=ft.border.all(1, color=st_color_Pink),
        content=ft.Column(
            controls=[
                BTN_Master_bat,
                BTN_Planner_bat,
                BTN_Hotkey_rvt,
                BTN_Matrix,
                BTN_SunPath,
                URL_row
                #BTN_Manager_IDS
            ],
            spacing=5,
            alignment=ft.MainAxisAlignment.START,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        ),
        padding=5,
    )
    
    column_MENU = ft.Column(
        controls=[
            logo_block,
            column_buttons,
        ],
        width=250,
        expand=True,
        spacing=10,
        alignment=ft.MainAxisAlignment.START,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER
    )

    #UI Меню

    page.add(
        ft.Column(  
            expand=True,
            width=250,
            alignment=ft.MainAxisAlignment.START,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[   
                ft.Container( 
                    expand=True,
                    width=250,
                    bgcolor=st_color_WHITE,
                    border_radius=ft.border_radius.all(10),
                    border=ft.border.all(2, color=st_color_BerryPink),
                    padding=5,
                    content=column_MENU
                ),
            ]
        )
    )



    
ft.app(target=main,assets_dir="assets")
