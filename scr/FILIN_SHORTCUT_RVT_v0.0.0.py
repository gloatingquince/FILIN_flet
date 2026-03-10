import flet as ft
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os

st_color_BLACK="#000000"
st_color_BerryPink="#D0034E"
st_color_AcePink="#F40059"
st_color_PurePink="#FFEBF3"
st_color_BgGrey="#E0E3E7"
st_color_txtGrey="#060606"

class XMLEditorApp:
    def __init__(self):
        self.xml_data = []
        self.file_path = None
        self.list_view = ft.ListView(expand=True, spacing=5, padding=5, cache_extent=500) 
        self.save_btn = None
        self.page = None
        
        self.info_text = ft.Text(
            "Импортируйте XML файл для начала работы", 
            expand=True, 
            size=14,
            selectable=True, 
            max_lines=None, 
            no_wrap=True,
            color=st_color_BLACK
        )

        self.reserved_info_text = ft.Text(
            value="",
            size=12,
            expand=True,
            selectable=True,
            max_lines=None,
            no_wrap=False,
            color=st_color_AcePink,
        )

        self.duplicates_info_text = ft.Text(
            value="",
            size=12,
            expand=True,
            selectable=True,
            max_lines=None,
            no_wrap=False,
            color=st_color_BLACK,
        )

        self.search_field = ft.TextField(
            hint_text="Поиск по Command",
            on_change=self.search_xml,
            color=st_color_BLACK,
            border_color=ft.Colors.GREY_900,
            cursor_color=st_color_AcePink,
            focused_border_color=st_color_AcePink,
            border_radius=8,
            focused_border_width=1,
            text_size=12,
            text_vertical_align=5,
            width=150,
            height=30,
        ) 

    def main(self, page: ft.Page):
        self.page = page
        page.bgcolor = "#EEEEEE"
        page.title = "XML Shortcuts Editor"
        page.vertical_alignment = ft.MainAxisAlignment.CENTER
        page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        page.window.min_width = 800
        page.window.width = 800
        page.padding = 10

        self.file_picker = ft.FilePicker(on_result=self.on_file_picked)
        self.save_file_picker = ft.FilePicker(on_result=self.on_save_location_picked)
        page.overlay.extend([self.file_picker, self.save_file_picker])

        def import_btn_clicked(_):
            self.info_text.value="Загрузка"
            self.info_text.color=st_color_AcePink
            self.info_text.size=20
            self.info_text.update()
            self.file_picker.pick_files(
                allowed_extensions=["xml"],
                dialog_title="Выберите XML файл"
            )

        import_btn = ft.OutlinedButton(
            text="Импорт XML",
            height=30,
            on_click=import_btn_clicked,
            style=ft.ButtonStyle(
                bgcolor="#FFEBF3",
                color="#060606",
                shape=ft.RoundedRectangleBorder(radius=8),
                side=ft.BorderSide(width=1, color=st_color_AcePink),
            )
        )

        self.save_btn = ft.OutlinedButton(
            text="Сохранить XML",
            height=30,
            on_click=self.save_xml,
            style=ft.ButtonStyle(
                bgcolor="#FFEBF3", color="#060606",
                shape=ft.RoundedRectangleBorder(radius=8),
                side=ft.BorderSide(width=1, color=st_color_AcePink)
            )
        )

        capslock_translit_btn = ft.OutlinedButton(
            text="CapsLockTranslit",
            height=30,
            tooltip="Добавить CapsLock / Транслит комбинации",
            on_click=self.on_capslock_translit_click,
            style=ft.ButtonStyle(
                bgcolor="#FFEBF3",
                color="#060606",
                shape=ft.RoundedRectangleBorder(radius=8),
                side=ft.BorderSide(width=1, color=st_color_AcePink)
            )
        )

        menu_row = ft.Row([
            import_btn, self.save_btn, capslock_translit_btn, self.search_field, self.info_text,
        ], adaptive=True, expand=True, run_spacing=10, alignment=ft.MainAxisAlignment.START)

        page.add(
            ft.Container(padding=2, content=menu_row),
            ft.Divider(thickness=1, color=st_color_AcePink),
            ft.Row([
                ft.Container(content=self.list_view, expand=3),
                ft.VerticalDivider(width=1, color=st_color_AcePink),
                ft.Container(
                    content=ft.Column([
                        self.duplicates_info_text,
                        self.reserved_info_text,
                    ], spacing=5, alignment = ft.MainAxisAlignment.START),
                    expand=2,
                ),
            ], expand=True)
        )

    def on_file_picked(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.file_path = e.files[0].path
            self.load_xml(self.file_path)
            self.validate_shortcuts()
        else:
            self.info_text.value=""
            self.info_text.update()
            self.info_text.size=14


    def load_xml(self, file_path):
        self.list_view.clean()

        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            self.xml_data.clear()

            for item in root.findall('ShortcutItem'):
                command_name = item.get('CommandName', '')
                command_id = item.get('CommandId', '')
                shortcuts_str = item.get('Shortcuts', '')
                paths = item.get('Paths', '')
                reserved_com_name=["Выход из Revit", "Закрыть", "Копировать в буфер", "Вырезать в буфер", "Вставить", "Отменить", "Повторить","Создать; Проект","Создать; Файл Revit","Печать","Сохранить","Поиск в Диспетчере проекта","Справка","Обновить","Проверка орфографии","Динамический вид","Переключить на главную", "Увеличить масштаб спецификации", "Уменьшить масштаб спецификации", "Восстановление масштаба вида спецификации",]                                                              

                shortcuts_field = ft.TextField(
                    height=35,
                    color=st_color_BLACK,
                    focused_border_width=1,
                    focused_border_color=st_color_AcePink,
                    border_radius=8,
                    cursor_color=st_color_AcePink,
                    text_size=12,
                    value=shortcuts_str,
                    filled=True,
                    bgcolor="#F8F7F8",
                    expand=True,
                    hint_text="Введите shortcuts через '#'",
                    read_only=False,
                    on_change=lambda e, item_data=None: self.on_shortcuts_changed(e, item_data)
                )
                
                prefix_buttons = ft.Row([
                    ft.OutlinedButton(
                        text="#",
                        tooltip="Добавить '#'",
                        style=ft.ButtonStyle(
                            bgcolor=None,
                            color="#36618E",
                            shape=ft.RoundedRectangleBorder(radius=8),
                            side=ft.BorderSide(width=1, color=st_color_AcePink)
                        ),
                        on_click=lambda e, f=shortcuts_field: self.add_prefix(f, "#")),
                    ft.OutlinedButton(
                        text="Ctrl+",
                        tooltip="Добавить 'Ctrl+'",
                        style=ft.ButtonStyle(
                            bgcolor=None,
                            color="#36618E",
                            shape=ft.RoundedRectangleBorder(radius=8),
                            side=ft.BorderSide(width=1, color=st_color_AcePink)
                        ),                        
                        on_click=lambda e, f=shortcuts_field: self.add_prefix(f, "Ctrl+")),
                    ft.OutlinedButton(
                        text="Shift+",
                        tooltip="Добавить 'Shift+'",
                        style=ft.ButtonStyle(
                            bgcolor=None,
                            color="#36618E",
                            shape=ft.RoundedRectangleBorder(radius=8),
                            side=ft.BorderSide(width=1, color=st_color_AcePink)
                        ),
                        on_click=lambda e, f=shortcuts_field: self.add_prefix(f, "Shift+")),
                ], disabled=False, spacing=5)

                item_data = {
                    'CommandName': command_name,
                    'CommandId': command_id,
                    'Paths': paths,
                    'shortcuts_field': shortcuts_field,
                    'card': None
                }
                self.xml_data.append(item_data)

                card = ft.Container(
                    content=ft.Column([
                        ft.Row([
                            ft.Text(f"Command: {command_name}", size=12, weight=ft.FontWeight.BOLD),
                            ft.Text(f"ID: {command_id}", size=10, color=ft.Colors.GREY_700),
                        ]),
                        ft.Row([shortcuts_field, prefix_buttons], alignment=ft.MainAxisAlignment.START),
                        ft.Text(f"Paths: {paths}", size=10, color=ft.Colors.GREY_600)
                    ]),
                    padding=5,
                    bgcolor=None
                )

                if command_name in reserved_com_name:
                    shortcuts_field.hint_text="Невозможно редактировать-зарезервированно системой"
                    card.bgcolor="#FFF0F6"
                    card.disabled=True
                else:
                    card.disabled=False

                item_data['card'] = card
                self.list_view.controls.append(card)

            self.page.update()
            self.info_text.value = ""
            self.info_text.size=14
            self.search_xml()
            self.info_text.update()
            

        except FileNotFoundError:
            self.show_error(f"Файл не найден по пути: {file_path}")
        except ET.ParseError:
            self.show_error("Ошибка парсинга XML файла. Убедитесь, что файл корректен.")
        except Exception as e:
            self.show_error(f"Произошла непредвиденная ошибка при загрузке: {str(e)}")

    def add_prefix(self, field: ft.TextField, prefix: str):
        current_value = field.value or ""
        cursor_pos = len(current_value)
        field.value = current_value[:cursor_pos] + prefix + current_value[cursor_pos:]
        self.page.update()

    def on_capslock_translit_click(self, e):
        translit_map = {
            "q": "й", "w": "ц", "e": "у", "r": "к", "t": "е", "y": "н", "u": "г", "i": "ш", "o": "щ", "p": "з",
            "@": "х", "@": "ъ", "a": "ф", "s": "ы", "d": "в", "f": "а", "g": "п", "h": "р", "j": "о", "k": "л",
            "l": "д", "@": "ж", "@": "э", "z": "я", "x": "ч", "c": "с", "v": "м", "b": "и", "n": "т", "m": "ь",
            "@": "б", "@": "ю"
        }

        def is_function_key(s):
            return any(
                part.upper().startswith("F") and part[1:].isdigit()
                for part in s.replace('FN', 'F').split('+')
            )

        def transliterate(text: str):
            return ''.join(
                translit_map.get(ch.lower(), ch).upper() if ch.isupper() else translit_map.get(ch.lower(), ch)
                for ch in text
            )

        for item in self.xml_data:
            field = item['shortcuts_field']
            current_raw = [s.strip() for s in field.value.split('#') if s.strip()]
            existing = set(current_raw)
            additions = set()

            for shortcut in current_raw:
                if is_function_key(shortcut):
                    continue

                if '+' in shortcut:
                    parts = shortcut.split('+')
                    prefix = '+'.join(parts[:-1])
                    key = parts[-1]

                    # Добавляем варианты регистра
                    for variant in {key.lower(), key.upper()}:
                        new_combo = f"{prefix}+{variant}"
                        if new_combo not in existing and new_combo not in additions:
                            additions.add(new_combo)

                    # Транслитерация только если key — английская
                    if all(c.lower() in translit_map for c in key):
                        translit = transliterate(key)
                        for t_variant in {translit.lower(), translit.upper()}:
                            full = f"{prefix}+{t_variant}"
                            if full not in existing and f"#{full}" not in existing and f"#{full}" not in additions:
                                additions.add(f"#{full}")

                else:
                    # Без '+', обычная комбинация
                    for variant in {shortcut.lower(), shortcut.upper()}:
                        if variant not in existing and variant not in additions:
                            additions.add(variant)

                    # Транслитерация (если англ)
                    if all(c.lower() in translit_map for c in shortcut):
                        translit = transliterate(shortcut)
                        for t_variant in {translit.lower(), translit.upper()}:
                            if t_variant not in existing and f"#{t_variant}" not in existing and f"#{t_variant}" not in additions:
                                additions.add(f"#{t_variant}")

            final = sorted(existing.union(additions))
            field.value = "#".join(final)
        
        self.validate_shortcuts()
        self.page.update()


    def on_shortcuts_changed(self, e, item_data):
        self.validate_shortcuts()

    def save_xml(self, e):
        if not self.xml_data:
            self.show_error("Нет данных для сохранения")
            return

        self.validate_shortcuts()  # Дубликаты не блокируют сохранение

        self.save_file_picker.save_file(
            dialog_title="Сохранить XML файл",
            file_name="shortcuts_edited.xml",
            allowed_extensions=["xml"]
        )

    def validate_shortcuts(self):
        all_shortcuts = {}
        has_duplicates = False
        reserved_warnings = {}
        reserved_shortcuts = set(s.upper() for s in [
        "ALT+FN4", "ALT+F4", "CTRL+W", "CTRL+C", "CTRL+INSERT", "CTRL+X", "SHIFT+DELETE",
        "CTRL+V", "CTRL+Z", "ALT+BACKSPACE", "CTRL+SHIFT+Z", "CTRL+Y", "CTRL+N", "CTRL+O",
        "CTRL+P", "CTRL+S", "CTRL+F", "F1/FN1", "F5/FN5", "F7/FN7", "F8/FN8", "SHIFT+W",
        "CTRL+D", "CTRL+F4", "TAB", "SHIFT+TAB", "ESC", "F10/FN10", "ENTER", "CTRL+0",
        "CTRL+(+)", "CTRL+(-)"
        ])

        for item in self.xml_data:
            command_name = item['CommandName']
            shortcuts_str = item['shortcuts_field'].value or ""
            current_shortcuts = [s.strip() for s in shortcuts_str.split('#') if s.strip()]

            for shortcut in current_shortcuts:
                upper_shortcut = shortcut.upper()
                if shortcut in all_shortcuts:
                    all_shortcuts[shortcut].append(f"{command_name} (ID: {item['CommandId']})")
                    has_duplicates = True
                else:
                    all_shortcuts[shortcut] = [f"{command_name} (ID: {item['CommandId']})"]

                if upper_shortcut in reserved_shortcuts:
                    if command_name not in reserved_warnings:
                        reserved_warnings[command_name] = []
                    reserved_warnings[command_name].append(shortcut)

        if reserved_warnings:
            msg = "Введённые комбинации зарезервированы системой, вы не можете их использовать:\n\n"
            for cmd, shortcuts in reserved_warnings.items():
                msg += f"{cmd}:\n"
                for s in shortcuts:
                    msg += f" - {s}\n"
                msg += "\n"
            self.reserved_info_text.value = msg
        else:
            self.reserved_info_text.value = ""

        if has_duplicates:
            msg = " Найдены дублирующиеся комбинации Shortcuts:\n\n  Дублирование разрешено. Перемещение по отображаемым в строке состояния записям выполняется с помощью клавиш со стрелками, выбор необходимой записи с помощью клавиши пробела.\n\n"
            for shortcut, sources in all_shortcuts.items():
                if len(sources) > 1:
                    msg += f"'{shortcut}' используется в:\n"
                    for src in sources:
                        msg += f"- {src}\n"
                    msg += "\n"
            self.duplicates_info_text.value = msg
        else:
            self.duplicates_info_text.value = "Дубликаты не выявлены"

        self.page.update()
        return True  # не блокирует экспорт

    def on_save_location_picked(self, e: ft.FilePickerResultEvent):
        if e.path:
            self.write_xml(e.path)

    def write_xml(self, file_path):
        try:
            root = ET.Element("Shortcuts")
            for item in self.xml_data:
                el = ET.SubElement(root, "ShortcutItem")
                el.set("CommandName", item['CommandName'])
                el.set("CommandId", item['CommandId'])
                el.set("Shortcuts", item['shortcuts_field'].value or "")
                el.set("Paths", item['Paths'])

            rough = ET.tostring(root, 'utf-8')
            parsed = minidom.parseString(rough)
            xml_str = parsed.toprettyxml(indent="  ", encoding="utf-8").decode('utf-8')
            xml_str = '\n'.join(xml_str.split('\n')[1:])

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(xml_str.strip())

            self.show_success(f"Файл успешно сохранён: {os.path.basename(file_path)}")

        except Exception as e:
            self.show_error(f"Ошибка при сохранении: {str(e)}")

    def show_error(self, message):
        self.info_text.value = message
        dlg = ft.AlertDialog(
            title=ft.Text("Ошибка"),
            content=ft.Text(message),
            actions=[ft.TextButton("OK", on_click=lambda _: self.close_dialog(dlg))],
            modal=True
        )
        self.page.dialog = dlg
        dlg.open = True
        self.page.update()

    def show_success(self, message):
        dlg = ft.AlertDialog(
            title=ft.Text("Успех"),
            content=ft.Text(message),
            actions=[ft.TextButton("OK", on_click=lambda _: self.close_dialog(dlg))],
            modal=True
        )
        self.page.dialog = dlg
        dlg.open = True
        self.page.update()

    def close_dialog(self, dlg):
        dlg.open = False
        self.page.dialog = None
        self.page.update()

    def search_xml(self, e=None):
        search_term = self.search_field.value.lower()
        filtered_cards = []

        for item in self.xml_data:
            if search_term in item['CommandName'].lower() or search_term in item['Paths'].lower():
                filtered_cards.append(item['card'])

        self.list_view.controls.clear()
        self.list_view.controls.extend(filtered_cards)
        self.page.update()

if __name__ == "__main__":
    app = XMLEditorApp()
    ft.app(target=app.main)
