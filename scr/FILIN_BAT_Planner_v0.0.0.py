import flet as ft
import os
import win32com.client
import datetime
import subprocess
import pythoncom

def main(page: ft.Page):
    page.title = "Планировщик Bat-файлов"
    page.bgcolor = "#EEEEEE"
    page.vertical_alignment = ft.MainAxisAlignment.START
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window.min_width = 600
    page.window.max_width = 600
    page.window.width = 600
    page.window.min_height = 780
    page.window.max_height = 800
    page.window.height = 780
    page.window.resizable=False
    page.window.always_on_top=True

    #Создаем стили
    #ST0 Цвета
    st_color_BLACK=ft.Colors.BLACK
    st_color_WHITE=ft.Colors.WHITE
    st_color_GREY=ft.Colors.ON_SURFACE_VARIANT
    st_color_AcePink=ft.Colors="#F40059"
    st_color_PurePink=ft.Colors="#FFEBF3" 
    st_color_AcentPink=ft.Colors="#E6A8B5"

    st_BTN_AllertDi=ft.ButtonStyle(
        color=st_color_AcePink,
    )
    
    #ST0 Стили текста
    st_main_name=ft.TextStyle(
        color=st_color_BLACK,
        font_family="Inter",
        weight=ft.FontWeight.BOLD,
        size=18,
    )

    st_second_name=ft.TextStyle(
        color=st_color_BLACK,
        weight=ft.FontWeight.BOLD,
        size=18
    )
    
    st_bat_name=ft.TextStyle(
        color=st_color_BLACK,
        weight=ft.FontWeight.BOLD,
        size=16,
    )

    st_txt_entry=ft.TextStyle(
        color=st_color_BLACK,
        size=16,
    )

    st_bat_path=ft.TextStyle(
        italic=True,
        color=st_color_GREY,
        size=12
    )
    st_bat_info=ft.TextStyle(
        color=st_color_GREY,
        size=12
    )
    
    #ST1 Стиль Кнопки

    BTN_AcePink=ft.ButtonStyle(
        bgcolor=st_color_PurePink,
        color=st_color_BLACK,
        shape=ft.RoundedRectangleBorder(radius=8),
        side=ft.BorderSide(width=1, color=st_color_AcePink)
    )

    found_bat_files_list = []

    def check_task_scheduler(file_path):
        pythoncom.CoInitialize()
        scheduler = win32com.client.Dispatch("Schedule.Service")
        scheduler.Connect()
        root_folder = scheduler.GetFolder("\\")
        try:
            trigger_types = {
                1: "Однократно", 2: "Ежедневно", 3: "Еженедельно", 4: "Ежемесячно", 
                5: "При входе", 6: "При простое", 7: "При регистрации", 8: "При загрузке", 
                9: "При входе в систему", 11: "Изменение состояния сеанса"
            }
            action_types={
                0: "Запустить программу", 5: "Вызвать обработчик",
                6: "Отправить сообщение электронной почты", 7: "Вывести сообщение"
            }

            for task in root_folder.GetTasks(0):
                if file_path in str(task.Definition.Actions.Item(1).Path):
                    triggers = []
                    for trigger in task.Definition.Triggers:
                        trigger_type = trigger_types.get(trigger.Type, "Другой")
                        triggers.append(f" Повтор: {trigger_type}, Начало: {trigger.StartBoundary}")
                    actions = []
                    for action in task.Definition.Actions:
                        action_type = action_types.get(action.Type, "Другой")
                        actions.append(f" \nТип: {action_type}\n  Путь: {action.Path}" )
                    return task.Name, "; ".join(triggers), "; ".join(actions)
            return None
        finally:
            pythoncom.CoUninitialize()
    
    
    def find_and_display_bat_files(e):
        selected_bats_path_input.value =(
        (e.path) if e.path else "Пути не получены((")
        selected_bats_path_input.update()
        if not selected_bats_path_input.value or not os.path.isdir(selected_bats_path_input.value):
            found_bat_files_list.clear()
            update_cards_display()
            return
        found_bat_files_list.clear()
        for root, _, files in os.walk(selected_bats_path_input.value):
            for file in files:
                if file.lower().endswith(".bat"):
                    full_path = os.path.join(root, file)
                    found_bat_files_list.append((file, full_path))
        update_cards_display()

    def update_cards_display():
        cards_column.controls.clear()

        if not found_bat_files_list:
            cards_column.controls.append(
                ft.Text("Файлы .bat не найдены или путь не указан.", 
                italic=True, 
                color=st_color_GREY,
                )
            )
        else:
            for filename, filepath in found_bat_files_list:
                task_info = check_task_scheduler(filepath)
                if task_info:
                    task_name, triggers, actions = task_info
                    task_card = ft.Container(
                        content=ft.Column([
                            ft.Text(f"Имя файла: {filename}", style=st_bat_name, selectable=True),
                            ft.Text(f"Путь: {filepath}", style=st_bat_path, selectable=True),
                            ft.Text(f"Задача: {task_name}", style=st_bat_info, selectable=True),
                            ft.Text(f"Триггеры: {triggers}", style=st_bat_info, selectable=True),
                            ft.Text(f"Действия: {actions}", style=st_bat_info, selectable=True),
                        ],
                        spacing=5),
                        alignment=ft.alignment.center,
                        border_radius=ft.border_radius.all(8),
                        border=ft.border.all(1, color=st_color_AcentPink),
                        padding=5
                    )
                else:
                    # Карточка для файла без задачи
                    task_card = ft.Container(
                        padding=5,
                        alignment=ft.alignment.center,
                        border_radius=ft.border_radius.all(8),
                        border=ft.border.all(1, color=st_color_AcentPink),
                        content=ft.Column([
                            ft.Text(f"Имя файла: {filename}", style=st_bat_name, selectable=True),
                            ft.Text(f"Путь: {filepath}", style=st_bat_path, selectable=True),
                            ft.OutlinedButton("Настроить и сохранить задачу", 
                                style=BTN_AcePink,  
                                on_click=lambda _, p=filepath: create_task_dialog(p)
                            ),
                        ],
                        spacing=5),
                    )
                cards_column.controls.append(task_card)
        page.update()

    def _close_dialog(e):
        """Закрывает диалоговое окно."""
        task_config_dialog.open = False
        page.update()

    def create_task_dialog(filepath_to_configure):
        """
        Открывает диалоговое окно для настройки параметров задачи.
        """
        task_name_field.value = f"BatTask_{os.path.basename(filepath_to_configure).replace('.', '_')}"
        task_time_field.value = "00:00"
        trigger_type_dropdown.value = "Однократно"
        day_of_week_dropdown.visible = False

        def _save_configured_task(e):
            task_name = task_name_field.value.strip()
            trigger_type_str = trigger_type_dropdown.value
            start_time_str = task_time_field.value.strip()
            selected_day = day_of_week_dropdown.value

            try:
                hour, minute = map(int, start_time_str.split(':'))
                if not (0 <= hour <= 23 and 0 <= minute <= 59):
                    raise ValueError("Неверный формат времени.")

                now = datetime.datetime.now()
                start_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                if start_time < now:
                    start_time += datetime.timedelta(days=1)

                start_boundary = start_time.strftime("%H:%M")
            except Exception as err:
                page.snack_bar = ft.SnackBar(ft.Text("Ошибка: время должно быть в формате ЧЧ:ММ"), open=True)
                task_config_dialog.open = False
                page.update()
                return

            schedule_map = {
                "Однократно": "ONCE",
                "Ежедневно": "DAILY",
                "Еженедельно": "WEEKLY",
            }
            schedule = schedule_map.get(trigger_type_str, "ONCE")

            cmd = [
                "schtasks",
                "/Create",
                "/SC", schedule,
                "/TN", task_name,
                "/TR", f'\"{filepath_to_configure}\"',
                "/ST", start_boundary,
                "/F"
            ]

            if schedule == "WEEKLY":
                day_map = {
                    "Понедельник": "MON",
                    "Вторник": "TUE",
                    "Среда": "WED",
                    "Четверг": "THU",
                    "Пятница": "FRI",
                    "Суббота": "SAT",
                    "Воскресенье": "SUN",
                }
                cmd.extend(["/D", day_map.get(selected_day, "MON")])

            try:
                result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
                if result.returncode == 0:
                    page.snack_bar = ft.SnackBar(ft.Text(f"Задача '{task_name}' успешно создана."), open=True)
                else:
                    raise Exception(result.stderr.strip())
            except Exception as ex:
                page.snack_bar = ft.SnackBar(ft.Text(f"Ошибка: {ex}"), open=True)

            task_config_dialog.open = False
            page.update()
            update_cards_display()

        task_config_dialog.actions[0].on_click = _save_configured_task

        page.dialog = task_config_dialog
        task_config_dialog.open = True
        page.update()
      
    def open_task_scheduler_gui():
        os.system("taskschd.msc")

    # Обработчик изменения типа запуска для видимости дня недели
    def _handle_trigger_type_change(e):
        day_of_week_dropdown.visible = trigger_type_dropdown.value == "Еженедельно"
        page.update()

    # UI BTN0 Кнопка поиска bat в директории
    search_bats_path = ft.FilePicker(on_result=find_and_display_bat_files)
    selected_bats_path_input = ft.TextField(
        text_style=st_txt_entry,
        border_color=st_color_BLACK,
        color=st_color_BLACK,
        value="Путь к директории для поиска .bat файлов",
        cursor_color=st_color_AcePink,
        focused_border_color=st_color_AcePink,
        filled=True,
        border_radius=8,
        focused_border_width=1,
        bgcolor="white",
        expand=True
        )

    page.overlay.append(search_bats_path)
  
    find_button = ft.OutlinedButton(
        style=BTN_AcePink,
        height=50, text="Выбрать",
        on_click=lambda _:search_bats_path.get_directory_path(dialog_title="Поиск Bat файлов")
        )
    
    # UI1 Карточки найденных батников
    cards_column = ft.Column(scroll=ft.ScrollMode.HIDDEN, spacing=10, expand=True)

    # UI1 Блок 1
    open_sched=ft.OutlinedButton(
        "Открыть планировщик задач",
        style=BTN_AcePink,
        on_click=lambda _: open_task_scheduler_gui()
    )
    
    search_block=ft.Column(
        spacing=15,
        controls=[
            ft.Row([
                ft.Text(
                    "Планировщик BAT",
                    style=st_main_name,
                    ),
                open_sched 
                ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            ),
            ft.Row(
               alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
               controls=(selected_bats_path_input, find_button), 
            )
        ],
    )

    # UI1 Блок 2
    cards_block=ft.Container(
        bgcolor=st_color_WHITE, 
        width = 600,
        height=800,
        padding=10,
        expand=True, 
        border_radius=ft.border_radius.all(8),
        border=ft.border.all(1, color=st_color_AcePink),
        content=ft.Column(
            expand=True,
            spacing=10,
            controls=[
                ft.Text("Найденные файлы:", style=st_second_name),
                ft.Divider(height=1, thickness=1, color=st_color_AcePink),
                ft.Container(
                    width = 600,
                    height=800,
                    expand=True,
                    alignment=ft.alignment.center,
                    content=cards_column,
                )
            ]
        )
    )

    # Диалоговое окно для настройки задачи
    task_name_field = ft.TextField(label="Имя задачи", value="", text_style=st_txt_entry, focused_border_color=st_color_AcePink,cursor_color=st_color_AcePink,)

    # Dropdown для выбора типа запуска
    trigger_type_dropdown = ft.Dropdown(
        text_style=st_txt_entry,
        label="Тип запуска",
        options=[
            ft.dropdown.Option("Однократно"),
            ft.dropdown.Option("Ежедневно"),
            ft.dropdown.Option("Еженедельно"),
        ],
        value="Однократно", # Значение по умолчанию
        on_change=_handle_trigger_type_change # Используем новую функцию-обработчик
    )

    # Dropdown для выбора дня недели (изначально скрыт)
    day_of_week_dropdown = ft.Dropdown(
        text_style=st_txt_entry,
        label="День недели",
        options=[
            ft.dropdown.Option("Понедельник"),
            ft.dropdown.Option("Вторник"),
            ft.dropdown.Option("Среда"),
            ft.dropdown.Option("Четверг"),
            ft.dropdown.Option("Пятница"),
            ft.dropdown.Option("Суббота"),
            ft.dropdown.Option("Воскресенье"),
        ],
        value="Понедельник", # Значение по умолчанию
        visible=False # Изначально скрыт
    )

    task_time_field = ft.TextField(
        text_style=st_txt_entry,
        label="Время запуска (ЧЧ:ММ)",
        value="00:00",
        hint_text="Например, 14:30",
        focused_border_color=st_color_AcePink,
        cursor_color=st_color_AcePink,
        )

    task_config_dialog = ft.AlertDialog(
        modal=True,
        bgcolor=st_color_WHITE,
        title=ft.Text("Настройка задачи", style=st_second_name),
        content=ft.Column([
            task_name_field,
            trigger_type_dropdown,
            day_of_week_dropdown,
            task_time_field,
        ], spacing=10),
        actions=[
            ft.TextButton("Сохранить",style=st_BTN_AllertDi, on_click=None), # on_click будет установлен динамически
            ft.TextButton("Отмена", style=st_BTN_AllertDi, on_click=_close_dialog),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )
    
    page.overlay.append(task_config_dialog) # Добавляем диалог в overlay страницы

    update_cards_display() # Инициализация отображения карточек

    page.add(
        ft.Container(
            expand=True,
            padding=5,
            content=ft.Column(
                expand=True,
                spacing=15,
                controls=[
                    search_block,
                    cards_block,
                ]
            )
        )
    )

    page.update()

ft.app(target=main)
