import pandas as pd
import json
import os
from datetime import datetime, timedelta
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from tkinter.scrolledtext import ScrolledText
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import re
import shutil
import webbrowser
import tempfile

# Проверка наличия reportlab
try:
    from reportlab.lib.pagesizes import A3, A4, landscape, portrait
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import cm, mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
    import reportlab.rl_config

    REPORTLAB_AVAILABLE = True

    try:
        font_paths = [
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/ariali.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
            "C:/Windows/Fonts/times.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/System/Library/Fonts/Helvetica.ttc"
        ]
        font_found = False
        for font_path in font_paths:
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont('Arial', font_path))
                font_found = True
                break
        if not font_found:
            reportlab.rl_config.warnOnMissingFontGlyphs = 0
    except:
        pass
except ImportError:
    REPORTLAB_AVAILABLE = False

# Проверка наличия matplotlib
try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure

    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False


class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Управление расписанием")
        self.root.geometry("1600x900")

        self.data_file = "schedule_data.xlsx"
        self.backup_folder = "backups"
        self.imported_files = []

        if not os.path.exists(self.backup_folder):
            os.makedirs(self.backup_folder)

        self.schedule_data = {}
        self.group_names = []
        self.days_of_week = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']
        self.current_day = self.days_of_week[0]
        self.current_group = None
        self.max_para = 6

        self.history_file = "schedule_history.json"
        self.history_data = self.load_history()

        self.is_dark_theme = False

        self.teachers_list = self.load_teachers_list()
        self.subjects_list = self.load_subjects_list()

        self.load_or_create_data()
        self.create_widgets()
        self.update_schedule_display()
        self.check_conflicts()
        self.setup_hotkeys()

    def setup_hotkeys(self):
        self.root.bind('<Control-n>', lambda e: self.add_lesson())
        self.root.bind('<Control-s>', lambda e: self.save_current_schedule())
        self.root.bind('<Control-f>', lambda e: self.show_search_dialog())
        self.root.bind('<Delete>', lambda e: self.delete_selected_lesson())
        self.root.bind('<Control-z>', lambda e: self.undo_action())

    def load_teachers_list(self):
        return [
            "Агарина Елена Михайловна", "Алейникова Лариса Петровна", "Алексеева Мария Евгеньевна",
            "Алексенцева Елена Лендрушевна", "Андреева Марина Михайловна", "Андриенко Елена Сергеевна",
            "Анисимова Татьяна Валентиновна", "Астраханцева Алена Сергеевна", "Белова Лилия Аркадьевна",
            "Беляев Виктор Анатольевич", "Близневская Наталья Борисовна", "Болгова Светлана Николаевна",
            "Бондаренко Анна-Мария Константиновна", "Бордунов Виктор Евгеньевич", "Бордунова Ольга Владиировна",
            "Борздая Алёна Николаевна", "Боярина Светлана Владимировна", "Быстрова Анна Дмитриевна",
            "Венидиктова Татьяна Ивановна", "Веселовская Екатерина Юрьевна", "Воевода Дарья Андреевна",
            "Волобуев Никита Юрьевич", "Выборкина Екатерина Михайловна", "Выходцева Наталья",
            "Гайдукова Ирина Леонидовна", "Галкин Руслан Александрович", "Глущенко Дмитрий Анатольевич",
            "Гончарова Вера Васильевна", "Горшков Евгений Викторович", "Горюнова Тамара Ивановна",
            "Грибанов Антон Владимирович", "Гуськова Юлия Ивановна", "Детиненко Татьяна Владимировна",
            "Джигирис Наталья Михайловна", "Доронина Елизавета Вячеславовна", "Егорова Светлана Викторовна",
            "Ерохина Ольга Михайловна", "Жабинцева Ирина Николаевна", "Замерец Ольга Васильевна",
            "Захарова Диана Вадимовна", "Захарова Светлана Викторовна", "Захарова-Ягодина Елена Аннатольевна",
            "Иващенко Ольга Сергеевна", "Ижевкина-Русинова Татьяна Святославовна", "Исакова Анастасия Дмитриевна",
            "Исакова Марина Викторовна", "Кайманакова Ольга Александровна", "Калмыкова Татьяна Анатольевна",
            "Карнаухова Татьяна Потаповна", "Карпегина Ольга Викторовна", "Ковалёв Артём Дмитриевич",
            "Кожанов Нурлан Тлепбайевич", "Кожевникова Анастасия Даниловна", "Козленя Анастасия Андреевна",
            "Козырева Милаида Сергеевна", "Колбина Юлия Леонидовна", "Конева Елена Васильевна",
            "Кононенко Виктор Владимирович", "Копытов Роман Борисович", "Корнева Алла Александровна",
            "Корнюхина Анастасия Игоревна", "Королева Елена Сергеевна", "Королёва Марина Владимировна",
            "Кравец Наталья Сергеевна", "Кравченко Екатерина Павловна", "Кузнецова Татьяна Юрьевна",
            "Кулакова Татьяна Сергеевна", "Лаврененкова Надежда Николаевна", "Лапина Лилия Владимировна",
            "Леонтьев Кирилл Александрович", "Лозбень Евгений Михайлович", "Лозбень Людмила Васильевна",
            "Любимова Марина Владимировна", "Ляхова Дарья Александровна", "Майснер Татьяна Васильевна",
            "Макарова Кристина Алексеевна", "Маниковская Елена Михайловна", "Медведева Екатерина Александровна",
            "Медведева Юлия Игоревна", "Мелентьев Кирилл Александрович", "Мелкова Татьяна Сергеевна",
            "Миллер Эдуард Александрович", "Михайлова Екатерина Владимировна", "Морокова Кристина Валерьевна",
            "Мурахтина Светлана Николаевна", "Мягкова Евгения Владимировна", "Неклюдов Иван Александрович",
            "Некрасова Ирина Игоревна", "Некрасова Светлана Александровна", "Немолочнова Светлана Анатольевна",
            "Никитина Анна Андреевна", "Новакович Евгения Викторовна", "Новикова Наталья Васильевна",
            "Огнев Семён Вячеславович", "Павлючик Сергей Константинович", "Пак Валентина Андреевна",
            "Паршакова Ольга Викторовна", "Паутова Лидия Ивановна", "Перебейнос Степан Григорьевич",
            "Плаксина Алла Анатольевна", "Побединский Александр Владимирович", "Повольских Яна Андреевна",
            "Попова Нина Михайловна", "Попрас Владимир Иванович", "Попрас Ирина Викторовна",
            "Потапов Роман Валерьевич", "Рубин Эвелина Витальевна", "Савин Игорь Владиславович",
            "Садыкова Евгения Владимировна", "Сайдаль Вячеслав Евгеньевич", "Салембаева Юлия Сергеевна",
            "Сальникова Ангелина Михайловна", "Самсонова Татьяна Владимировна", "Сартаков Алексей Андреевич",
            "Скальская Евгения Александровна", "Слепцова Анна Вячеславна", "Соболева Елена Викторовна",
            "Стародубцев Александр Генадьевич", "Суворова Наталья Владимировна", "Сысун Яна Сергеевна",
            "Тамарова Ирина Вячеславовна", "Тихомирова Татьяна Александровна", "Устюжанина Лилия Александровна",
            "Ушакова Наталья Владимировна", "Хабарова Наталья Викторовна", "Хабеев Илья Юрьевич",
            "Хахалина Анастасия Игоревна", "Чайковская Юлия Алексеевна", "Черных Сергей Николаевич",
            "Чоботова Светлана Михайловна", "Шелудько Ульяна Владимировна", "Ширшакова Елена Сергеевна",
            "Шишкина Ольга Николаевна", "Шмырева Елена Михайловна", "Шугаева Елена Владимировна",
            "Шуненко Наталья Михайловна", "Эйсайед Жанна Александровна", "Юзова Ольга Викторовна",
            "Языков Илья Александрович", "Яковлева Татьяна Анатольевна"
        ]

    def load_subjects_list(self):
        return [
            "Анализ", "Анализ муз. произв.", "БЖ", "Биология", "Возрастная психология*",
            "География", "Гармония", "Гармония (I)", "Гармония (I+II)",
            "ДР.МДК 02.01 (I)", "ДР.МДК 02.01 (II)", "ДР.МДК 02.01 ТИИ (II)",
            "ДР.МДК 02.01.(II)", "ДР.МДК 02.01.ТИИ (I)", "ДР.МДК 02.01.ТИИ (II)",
            "Живопись (I)/Рисунок (II)", "Инд.проект/Рус.яз.", "Информатика",
            "Информатика (1)*/Информатика (2)**", "Информатика I/МДК 02.01 ОБФиК II",
            "Информатика II/МДК 02.01 ОБФиК I", "Информат. (I)*/УП 02 (I)**",
            "Инф (I)*/Инф (II)**", "Иностранный язык", "Иностранный язык (2)",
            "Иностранный язык ОГСЭ", "Иностранный язык ОД", "ИОК", "История",
            "История / ОБиЗР", "История искусств**", "История исп.иск.", "История РФ",
            "Литература", "Математика", "Математика / Литература",
            "МДК 01.01 М.Р. Основы режиссуры", "МДК 01.01 Хоровое пение",
            "МДК 01.01 ХПИ (II)", "МДК 01.02 Библиографоведение**/МДК 02.01 Менеджмент*",
            "МДК 01.02 ИП Грим (1)/МДК 01.02 ИП Сц.движ (2)",
            "МДК 01.02 ИП ОАМ (1)/МДК 01.02 ИП Сц.речь (2)",
            "МДК 01.02 ИП Сц.речь (1)/МДК 01.02 ИП Танец (2)",
            "МДК 01.02 ИП Сц.речь (2)", "МДК 01.02 УП 00 исполнит.",
            "МДК 01.02 Хор.подгот.(класс.т)", "МДК 01.02 Хор.подгот.(лат)",
            "МДК 01.02 Хор.подгот.(совр.т)", "МДК 01.02 Хор.подгот.(станд)",
            "МДК 01.04 История исп.иск.", "МДК 02.01 (I)", "МДК 02.01 (II)",
            "МДК 02.01 Методика преп.хор.дисц.", "МДК 02.01 ОБФиК I/МДК 04.01 ИТПД II",
            "МДК 02.01 ОР и СМ РКМ и ТП", "МДК 02.01 ТИИ (I)/ДР.МДК 02.01 (II)",
            "МДК 02.01 ТИИ (II)", "МДК 02.01.ТИИ (I)", "МДК 02.01.ТИИ (II)",
            "МДК 02.02 ИП ОАМ (1)", "МДК 02.02 ИП ОАМ (2)", "МДК 03.01 ОДМ",
            "МДК 03.01 Хор.литература", "МДК 03.01 ЭиМ",
            "МДК 03.01 ЭиМ СКД/Основы финансовой грамотности",
            "Муз. литература", "Муз.подготовка", "Муз.литература",
            "Музыкальная литература", "Народный танец", "НМК",
            "Общее библиографоведение", "Обществознание", "ОП СКД",
            "Оркестр", "Основы психологии/психология общения",
            "Основы философии", "Основы филос.", "Оформление постановок",
            "Отечественная литература", "Рисунок (I)", "Рисунок (II)",
            "Родная литература", "Русский язык", "Русский язык и культура речи",
            "Сольфеджио", "Сольфеджио (I)", "УП 02 (I)",
            "УП 02.01 Игровые технологии", "УП 04 Ансамблевое пение",
            "УП ПМ.03", "УП.04 (I)", "Учеб.практика",
            "Учеб.практика 02.01 РКММ и ТП", "Учебн.практ (класс.т)",
            "Физ.культура", "Физика", "Физическая культура",
            "Фольк.ансамбль", "Химия", "Хор", "Хоровая литература",
            "Хороведение", "Экон.и мен.СКД**", "ЭТМ"
        ]

    def load_or_create_data(self):
        if os.path.exists(self.data_file):
            try:
                self.load_from_excel(self.data_file)
            except:
                self.load_excel_data()
        else:
            if os.path.exists('общий массив.xlsx'):
                self.load_excel_data()
            elif os.path.exists('общий массив — копия.xlsx'):
                self.load_excel_data('общий массив — копия.xlsx')
            else:
                self.create_test_data()
            self.save_to_excel(self.data_file)

    def load_excel_data(self, file_path='общий массив.xlsx'):
        try:
            if not os.path.exists(file_path):
                if file_path == 'общий массив.xlsx' and os.path.exists('общий массив — копия.xlsx'):
                    file_path = 'общий массив — копия.xlsx'
                else:
                    return False

            xl = pd.ExcelFile(file_path)
            self.schedule_data = {}
            self.group_names = []

            for sheet_name in xl.sheet_names:
                if sheet_name == 'Лист2':
                    continue

                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

                i = 0
                while i < len(df):
                    row = df.iloc[i]

                    if isinstance(row.iloc[0], str) and 'РАСПИСАНИЕ' in str(row.iloc[0]):
                        i += 1
                        if i >= len(df):
                            break

                        header_row = df.iloc[i]
                        if 'Группа' in str(header_row.iloc[0]) or 'группа' in str(header_row.iloc[0]):
                            i += 1
                            while i < len(df):
                                data_row = df.iloc[i]
                                if isinstance(data_row.iloc[0], str) and 'РАСПИСАНИЕ' in str(data_row.iloc[0]):
                                    break

                                group = str(data_row.iloc[0]) if not pd.isna(data_row.iloc[0]) else ''
                                day = str(data_row.iloc[1]) if not pd.isna(data_row.iloc[1]) else ''
                                para = str(data_row.iloc[2]) if not pd.isna(data_row.iloc[2]) else ''
                                subject = str(data_row.iloc[3]) if not pd.isna(data_row.iloc[3]) else ''
                                teacher = str(data_row.iloc[4]) if not pd.isna(data_row.iloc[4]) else ''
                                room = str(data_row.iloc[5]) if not pd.isna(data_row.iloc[5]) else ''

                                if group and day and para and subject and subject.strip():
                                    if group not in self.schedule_data:
                                        self.schedule_data[group] = {d: {} for d in self.days_of_week}

                                    day_normalized = self.normalize_day(day)
                                    if day_normalized:
                                        if day_normalized not in self.schedule_data[group]:
                                            self.schedule_data[group][day_normalized] = {}

                                        self.schedule_data[group][day_normalized][para] = {
                                            'предмет': subject,
                                            'преподаватель': teacher,
                                            'кабинет': room,
                                            'часы': 2
                                        }

                                i += 1
                            continue
                    i += 1

            self.group_names = sorted(self.schedule_data.keys())

            if self.group_names:
                self.current_group = self.group_names[0]
                return True
            else:
                self.create_test_data()
                return False

        except Exception as e:
            self.create_test_data()
            return False

    def normalize_day(self, day):
        day_mapping = {
            'Пн': 'Понедельник', 'Понедельник': 'Понедельник',
            'Вт': 'Вторник', 'Вторник': 'Вторник',
            'Ср': 'Среда', 'Среда': 'Среда',
            'Чт': 'Четверг', 'Четверг': 'Четверг',
            'Пт': 'Пятница', 'Пятница': 'Пятница',
            'Сб': 'Суббота', 'Суббота': 'Суббота'
        }
        return day_mapping.get(day, day)

    def load_from_excel(self, file_path):
        try:
            df = pd.read_excel(file_path, sheet_name='Данные', header=0)

            groups_df = pd.read_excel(file_path, sheet_name='Группы', header=0)
            self.group_names = groups_df['Группа'].tolist() if 'Группа' in groups_df.columns else []

            self.schedule_data = {}
            for group in self.group_names:
                self.schedule_data[group] = {day: {} for day in self.days_of_week}

            for _, row in df.iterrows():
                group = str(row['Группа']) if not pd.isna(row['Группа']) else ''
                day = str(row['День']) if not pd.isna(row['День']) else ''
                para = str(row['Пара']) if not pd.isna(row['Пара']) else ''
                subject = str(row['Предмет']) if not pd.isna(row['Предмет']) else ''
                teacher = str(row['Преподаватель']) if not pd.isna(row['Преподаватель']) else ''
                room = str(row['Кабинет']) if not pd.isna(row['Кабинет']) else ''
                hours = float(row['Часы']) if not pd.isna(row['Часы']) else 2

                if group and day and para and subject:
                    if group not in self.schedule_data:
                        self.schedule_data[group] = {d: {} for d in self.days_of_week}
                    day_norm = self.normalize_day(day)
                    if day_norm not in self.schedule_data[group]:
                        self.schedule_data[group][day_norm] = {}
                    self.schedule_data[group][day_norm][para] = {
                        'предмет': subject,
                        'преподаватель': teacher,
                        'кабинет': room,
                        'часы': hours
                    }

            if self.group_names:
                self.current_group = self.group_names[0]
            else:
                self.create_test_data()

        except Exception as e:
            raise e

    def save_to_excel(self, file_path=None):
        if file_path is None:
            file_path = self.data_file

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                rows = []
                for group in self.group_names:
                    for day in self.days_of_week:
                        for para, lesson in self.schedule_data.get(group, {}).get(day, {}).items():
                            rows.append({
                                'Группа': group,
                                'День': day,
                                'Пара': para,
                                'Предмет': lesson.get('предмет', ''),
                                'Преподаватель': lesson.get('преподаватель', ''),
                                'Кабинет': lesson.get('кабинет', ''),
                                'Часы': lesson.get('часы', 2)
                            })

                if rows:
                    df = pd.DataFrame(rows)
                    df.to_excel(writer, sheet_name='Данные', index=False)
                else:
                    pd.DataFrame(
                        columns=['Группа', 'День', 'Пара', 'Предмет', 'Преподаватель', 'Кабинет', 'Часы']).to_excel(
                        writer, sheet_name='Данные', index=False)

                pd.DataFrame({'Группа': self.group_names}).to_excel(writer, sheet_name='Группы', index=False)

                if self.imported_files:
                    pd.DataFrame({'Импортированные файлы': self.imported_files}).to_excel(
                        writer, sheet_name='Импорт', index=False)

            self.create_backup()
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить данные: {str(e)}")
            return False

    def create_backup(self):
        try:
            if os.path.exists(self.data_file):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_file = os.path.join(self.backup_folder, f"schedule_backup_{timestamp}.xlsx")
                shutil.copy2(self.data_file, backup_file)

                backups = sorted([f for f in os.listdir(self.backup_folder) if f.startswith('schedule_backup_')])
                if len(backups) > 10:
                    for old_backup in backups[:-10]:
                        os.remove(os.path.join(self.backup_folder, old_backup))
        except:
            pass

    def restore_backup(self):
        if not os.path.exists(self.backup_folder):
            messagebox.showwarning("Внимание", "Нет папки с бэкапами")
            return

        backups = sorted([f for f in os.listdir(self.backup_folder) if f.startswith('schedule_backup_')], reverse=True)
        if not backups:
            messagebox.showwarning("Внимание", "Нет бэкапов для восстановления")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Восстановление из бэкапа")
        dialog.geometry("600x450")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Выберите бэкап для восстановления:", font=('Arial', 12)).pack(pady=10)

        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        listbox = tk.Listbox(list_frame, height=15, width=70, font=('Arial', 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.config(yscrollcommand=scrollbar.set)

        for backup in backups:
            backup_path = os.path.join(self.backup_folder, backup)
            size = os.path.getsize(backup_path) / 1024
            date_str = backup.replace('schedule_backup_', '').replace('.xlsx', '')
            listbox.insert(tk.END, f"{date_str}  ({size:.1f} KB)")

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        def restore_selected():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Внимание", "Выберите бэкап")
                return

            backup_name = backups[selection[0]]
            backup_path = os.path.join(self.backup_folder, backup_name)

            if messagebox.askyesno("Подтверждение", f"Восстановить из бэкапа?\nТекущие данные будут потеряны."):
                try:
                    shutil.copy2(backup_path, self.data_file)
                    self.load_from_excel(self.data_file)
                    self.refresh_all()
                    dialog.destroy()
                    messagebox.showinfo("Успех", "Данные восстановлены из бэкапа")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось восстановить данные: {str(e)}")

        ttk.Button(btn_frame, text="Восстановить", command=restore_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

    def create_test_data(self):
        self.group_names = ['10', '20', '30', '13а']
        self.schedule_data = {}

        for group in self.group_names:
            self.schedule_data[group] = {day: {} for day in self.days_of_week}

        test_schedule = {
            '10': {
                'Понедельник': {
                    '1': {'предмет': 'Математика', 'преподаватель': 'Новикова Наталья Васильевна', 'кабинет': '121',
                          'часы': 2},
                    '2': {'предмет': 'Музыкальная литература', 'преподаватель': 'Медведева Екатерина Александровна',
                          'кабинет': '304', 'часы': 2},
                },
                'Вторник': {
                    '1': {'предмет': 'Химия', 'преподаватель': 'Жабинцева Ирина Николаевна', 'кабинет': '015(2)',
                          'часы': 2},
                }
            }
        }

        for group, days in test_schedule.items():
            if group in self.schedule_data:
                for day, paras in days.items():
                    if day in self.schedule_data[group]:
                        self.schedule_data[group][day].update(paras)

        if self.group_names:
            self.current_group = self.group_names[0]

    def load_history(self):
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_history(self):
        with open(self.history_file, 'w', encoding='utf-8') as f:
            json.dump(self.history_data, f, ensure_ascii=False, indent=2)

    def undo_action(self):
        self.restore_backup()

    def check_conflicts(self):
        conflicts = []

        teacher_schedule = defaultdict(list)
        for group in self.schedule_data:
            for day in self.days_of_week:
                for para, lesson in self.schedule_data.get(group, {}).get(day, {}).items():
                    teacher = lesson.get('преподаватель', '')
                    if teacher:
                        key = f"{day}_{para}"
                        for t in re.split(r'[/,]', teacher):
                            t = t.strip()
                            if t:
                                teacher_schedule[f"{t}_{key}"].append(group)

        for key, groups in teacher_schedule.items():
            if len(groups) > 1:
                parts = key.split('_')
                teacher = parts[0]
                day = parts[1]
                para = parts[2]
                conflicts.append(
                    f"Преподаватель {teacher} занят одновременно в группах {', '.join(groups)} на {day}, пара {para}")

        if conflicts:
            conflict_text = "\n".join(conflicts)
            if messagebox.askyesno("Обнаружены конфликты",
                                   f"Обнаружены следующие конфликты:\n\n{conflict_text}\n\nПоказать подробности?"):
                self.show_conflict_details(conflicts)

        return conflicts

    def show_conflict_details(self, conflicts):
        dialog = tk.Toplevel(self.root)
        dialog.title("Конфликты в расписании")
        dialog.geometry("700x450")
        dialog.transient(self.root)

        text = ScrolledText(dialog, width=80, height=20, font=('Arial', 10))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        text.insert(tk.END, "⚠️ ОБНАРУЖЕНЫ КОНФЛИКТЫ В РАСПИСАНИИ\n", 'header')
        text.insert(tk.END, "=" * 60 + "\n\n", 'header')

        for conflict in conflicts:
            text.insert(tk.END, f"• {conflict}\n", 'conflict')

        text.tag_config('header', foreground='red', font=('Arial', 12, 'bold'))
        text.tag_config('conflict', foreground='darkorange', font=('Arial', 10))
        text.config(state=tk.DISABLED)

        ttk.Button(dialog, text="Закрыть", command=dialog.destroy).pack(pady=10)

    def show_search_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Поиск")
        dialog.geometry("550x450")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Введите текст для поиска:", font=('Arial', 10)).pack(pady=5)

        search_var = tk.StringVar()
        search_entry = ttk.Entry(dialog, textvariable=search_var, width=50, font=('Arial', 10))
        search_entry.pack(pady=5)

        search_type = tk.StringVar(value="all")
        ttk.Label(dialog, text="Искать в:", font=('Arial', 10)).pack(pady=5)

        type_frame = ttk.Frame(dialog)
        type_frame.pack(pady=5)
        ttk.Radiobutton(type_frame, text="Везде", variable=search_type, value="all").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(type_frame, text="Предметах", variable=search_type, value="subject").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(type_frame, text="Преподавателях", variable=search_type, value="teacher").pack(side=tk.LEFT,
                                                                                                       padx=10)
        ttk.Radiobutton(type_frame, text="Кабинетах", variable=search_type, value="room").pack(side=tk.LEFT, padx=10)

        results_text = ScrolledText(dialog, height=12, width=70, font=('Arial', 10))
        results_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        def do_search():
            results_text.delete(1.0, tk.END)
            search_text = search_var.get().lower()
            if not search_text:
                messagebox.showwarning("Внимание", "Введите текст для поиска")
                return

            found = False

            for group in self.schedule_data:
                for day in self.days_of_week:
                    for para, lesson in self.schedule_data.get(group, {}).get(day, {}).items():
                        match = False
                        if search_type.get() in ['all', 'subject'] and search_text in lesson.get('предмет', '').lower():
                            match = True
                        if search_type.get() in ['all', 'teacher'] and search_text in lesson.get('преподаватель',
                                                                                                 '').lower():
                            match = True
                        if search_type.get() in ['all', 'room'] and search_text in lesson.get('кабинет', '').lower():
                            match = True

                        if match:
                            found = True
                            results_text.insert(tk.END,
                                                f"Группа: {group}, {day}, пара {para}\n"
                                                f"  Предмет: {lesson.get('предмет', '')}\n"
                                                f"  Преподаватель: {lesson.get('преподаватель', '')}\n"
                                                f"  Кабинет: {lesson.get('кабинет', '')}\n"
                                                f"  Часы: {lesson.get('часы', 2)}\n\n")

            if not found:
                results_text.insert(tk.END, "Ничего не найдено")

        search_entry.bind('<Return>', lambda e: do_search())

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Искать", command=do_search).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Закрыть", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

        search_entry.focus()

    def create_widgets(self):
        self.create_menu()

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(schedule_frame, text="📚 Расписание")
        self.create_schedule_tab(schedule_frame)

        editor_frame = ttk.Frame(self.notebook)
        self.notebook.add(editor_frame, text="✏️ Редактор")
        self.create_editor_tab(editor_frame)

        view_frame = ttk.Frame(self.notebook)
        self.notebook.add(view_frame, text="👁️ Просмотр")
        self.create_view_tab(view_frame)

        teachers_frame = ttk.Frame(self.notebook)
        self.notebook.add(teachers_frame, text="👨‍🏫 Преподаватели")
        self.create_teachers_tab(teachers_frame)

        rooms_frame = ttk.Frame(self.notebook)
        self.notebook.add(rooms_frame, text="🏢 Кабинеты")
        self.create_rooms_tab(rooms_frame)

        stats_frame = ttk.Frame(self.notebook)
        self.notebook.add(stats_frame, text="📊 Статистика")
        self.create_stats_tab(stats_frame)

        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="⚙️ Настройки")
        self.create_settings_tab(settings_frame)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Импорт из Excel", command=self.import_from_excel)
        file_menu.add_command(label="Экспорт в Excel", command=self.export_to_excel)
        if REPORTLAB_AVAILABLE:
            file_menu.add_command(label="Экспорт в PDF", command=self.export_to_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="Создать бэкап", command=self.create_backup_manual)
        file_menu.add_command(label="Восстановить из бэкапа", command=self.restore_backup)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.quit)

        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Редактирование", menu=edit_menu)
        edit_menu.add_command(label="Добавить пару (Ctrl+N)", command=self.add_lesson)
        edit_menu.add_command(label="Поиск (Ctrl+F)", command=self.show_search_dialog)
        edit_menu.add_command(label="Отмена (Ctrl+Z)", command=self.undo_action)

        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Вид", menu=view_menu)
        view_menu.add_command(label="Обновить", command=self.refresh_all)

        reports_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Отчеты", menu=reports_menu)
        reports_menu.add_command(label="Расписание на день (все группы)", command=self.show_all_groups_day)
        reports_menu.add_command(label="Расписание на неделю (все группы)", command=self.show_all_groups_week)
        reports_menu.add_command(label="Часы преподавателей", command=self.show_teachers_hours)
        reports_menu.add_command(label="Статистика по кабинетам", command=self.show_rooms_stats)
        reports_menu.add_command(label="Экспорт расписания на день", command=self.export_day_schedule)
        reports_menu.add_command(label="Экспорт расписания на неделю", command=self.export_week_schedule)
        reports_menu.add_command(label="Экспорт в Google Calendar", command=self.export_to_google_calendar)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Справка", menu=help_menu)
        help_menu.add_command(label="Горячие клавиши", command=self.show_shortcuts)
        help_menu.add_command(label="О программе", command=self.show_about)

    def create_schedule_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=5)

        ttk.Label(control_frame, text="Группа:").pack(side=tk.LEFT, padx=5)
        self.group_var = tk.StringVar()
        self.group_combo = ttk.Combobox(control_frame, textvariable=self.group_var,
                                        values=self.group_names, width=15, state='readonly')
        self.group_combo.pack(side=tk.LEFT, padx=5)
        self.group_combo.bind('<<ComboboxSelected>>', self.on_group_change)
        if self.current_group:
            self.group_var.set(self.current_group)

        ttk.Label(control_frame, text="День:").pack(side=tk.LEFT, padx=5)
        self.day_var = tk.StringVar()
        self.day_combo = ttk.Combobox(control_frame, textvariable=self.day_var,
                                      values=self.days_of_week, width=15, state='readonly')
        self.day_combo.pack(side=tk.LEFT, padx=5)
        self.day_var.set(self.current_day)
        self.day_combo.bind('<<ComboboxSelected>>', self.on_day_change)

        btn_frame = ttk.Frame(control_frame)
        btn_frame.pack(side=tk.LEFT, padx=10)

        ttk.Button(btn_frame, text="Добавить (Ctrl+N)", command=self.add_lesson).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Сохранить (Ctrl+S)", command=self.save_current_schedule).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Подсчет часов", command=self.show_hours_count).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🔍 Поиск", command=self.show_search_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📋 На день", command=self.export_day_schedule).pack(side=tk.LEFT, padx=2)

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        columns = ('Пара', 'Предмет', 'Преподаватель', 'Кабинет', 'Часы')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)

        self.tree.heading('Пара', text='Пара')
        self.tree.heading('Предмет', text='Предмет')
        self.tree.heading('Преподаватель', text='Преподаватель')
        self.tree.heading('Кабинет', text='Кабинет')
        self.tree.heading('Часы', text='Часы')

        self.tree.column('Пара', width=60)
        self.tree.column('Предмет', width=350)
        self.tree.column('Преподаватель', width=250)
        self.tree.column('Кабинет', width=120)
        self.tree.column('Часы', width=60)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.context_menu = tk.Menu(self.tree, tearoff=0)
        self.context_menu.add_command(label="Удалить пару (Delete)", command=self.delete_selected_lesson)
        self.context_menu.add_command(label="Редактировать пару", command=self.edit_selected_lesson)
        self.tree.bind("<Button-3>", self.show_context_menu)

        history_frame = ttk.LabelFrame(main_frame, text="История расписаний")
        history_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        history_btn_frame = ttk.Frame(history_frame)
        history_btn_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(history_btn_frame, text="Очистить историю", command=self.clear_history).pack(side=tk.LEFT, padx=2)
        ttk.Button(history_btn_frame, text="Экспорт истории", command=self.export_history).pack(side=tk.LEFT, padx=2)

        self.history_text = ScrolledText(history_frame, height=6, width=80, font=('Arial', 9))
        self.history_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.update_history_display()

    def create_editor_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=5)

        ttk.Label(control_frame, text="Группа:").pack(side=tk.LEFT, padx=5)
        self.editor_group_var = tk.StringVar()
        self.editor_group_combo = ttk.Combobox(control_frame, textvariable=self.editor_group_var,
                                               values=self.group_names, width=15, state='readonly')
        self.editor_group_combo.pack(side=tk.LEFT, padx=5)
        self.editor_group_combo.bind('<<ComboboxSelected>>', self.on_editor_group_change)

        ttk.Button(control_frame, text="Создать неделю", command=self.create_weekly_schedule).pack(side=tk.LEFT,
                                                                                                   padx=10)
        ttk.Button(control_frame, text="Очистить", command=self.clear_weekly_schedule).pack(side=tk.LEFT, padx=5)

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        columns = ['Пара'] + self.days_of_week
        self.editor_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=8)

        self.editor_tree.heading('Пара', text='Пара')
        for day in self.days_of_week:
            self.editor_tree.heading(day, text=day[:3])

        self.editor_tree.column('Пара', width=60)
        for day in self.days_of_week:
            self.editor_tree.column(day, width=180)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.editor_tree.yview)
        self.editor_tree.configure(yscrollcommand=scrollbar.set)

        self.editor_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.editor_tree.bind("<Double-1>", self.edit_cell)

        info_frame = ttk.LabelFrame(main_frame, text="Инструкция")
        info_frame.pack(fill=tk.X, pady=5)

        info_text = tk.Text(info_frame, height=3, width=80, state=tk.DISABLED, font=('Arial', 9))
        info_text.pack(fill=tk.X, padx=5, pady=5)
        info_text.config(state=tk.NORMAL)
        info_text.insert(1.0, "Двойной клик по ячейке для редактирования.\n"
                              "Формат: Предмет|Преподаватель|Кабинет|Часы\n"
                              "Пример: Математика|Новикова Н.В.|121|2")
        info_text.config(state=tk.DISABLED)

        ttk.Button(main_frame, text="💾 Сохранить", command=self.save_editor_schedule).pack(pady=10)

    def create_view_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=5)

        ttk.Label(control_frame, text="Режим просмотра:", font=('Arial', 10)).pack(side=tk.LEFT, padx=5)

        self.view_mode = tk.StringVar(value="day")
        ttk.Radiobutton(control_frame, text="День", variable=self.view_mode, value="day",
                        command=self.update_view).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(control_frame, text="Неделя", variable=self.view_mode, value="week",
                        command=self.update_view).pack(side=tk.LEFT, padx=5)

        ttk.Label(control_frame, text="День:").pack(side=tk.LEFT, padx=10)
        self.view_day_var = tk.StringVar()
        self.view_day_combo = ttk.Combobox(control_frame, textvariable=self.view_day_var,
                                           values=self.days_of_week, width=15, state='readonly')
        self.view_day_combo.pack(side=tk.LEFT, padx=5)
        self.view_day_var.set(self.days_of_week[0])
        self.view_day_combo.bind('<<ComboboxSelected>>', lambda e: self.update_view())

        ttk.Button(control_frame, text="Экспорт в Excel", command=self.export_view_to_excel).pack(side=tk.LEFT, padx=5)
        if REPORTLAB_AVAILABLE:
            ttk.Button(control_frame, text="Экспорт в PDF", command=self.export_view_to_pdf).pack(side=tk.LEFT, padx=5)

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.view_tree = ttk.Treeview(table_frame, show='headings', height=20)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.view_tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.view_tree.xview)
        self.view_tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.view_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        info_frame = ttk.LabelFrame(main_frame, text="Информация")
        info_frame.pack(fill=tk.X, pady=5)

        self.view_info_text = tk.Text(info_frame, height=2, width=80, state=tk.DISABLED, font=('Arial', 9))
        self.view_info_text.pack(fill=tk.X, padx=5, pady=5)

        self.update_view()

    def update_view(self):
        mode = self.view_mode.get()
        if mode == "day":
            self.show_day_view()
        else:
            self.show_week_view()

    def show_day_view(self):
        """Показ расписания на день для всех групп"""
        for item in self.view_tree.get_children():
            self.view_tree.delete(item)

        day = self.view_day_var.get()
        if not day:
            return

        # Настраиваем колонки - убираем колонку "Часы"
        columns = ['Группа'] + [f'{i} пара' for i in range(1, self.max_para + 1)] + ['Итого часов']
        self.view_tree['columns'] = columns

        for col in columns:
            self.view_tree.heading(col, text=col)
            if col == 'Группа':
                self.view_tree.column(col, width=100)
            elif col == 'Итого часов':
                self.view_tree.column(col, width=100)
            else:
                self.view_tree.column(col, width=180)

        # Заполняем данными
        for group in sorted(self.group_names):
            day_schedule = self.schedule_data.get(group, {}).get(day, {})
            row = [group]
            total_hours = 0

            for para in range(1, self.max_para + 1):
                para_str = str(para)
                lesson = day_schedule.get(para_str, {})
                if lesson:
                    # В ячейке отображаем: предмет, преподаватель, кабинет (без часов)
                    text = f"{lesson.get('предмет', '')}\n{self.shorten_teacher_name(lesson.get('преподаватель', ''))}\n{lesson.get('кабинет', '')}"
                    hours = lesson.get('часы', 2)
                    total_hours += hours
                else:
                    text = ""
                row.append(text)

            row.append(str(total_hours))
            self.view_tree.insert('', tk.END, values=row)

        # Обновляем информацию
        self.view_info_text.config(state=tk.NORMAL)
        self.view_info_text.delete(1.0, tk.END)
        info_text = f"📅 Расписание на {day}\n"
        info_text += f"📚 Всего групп: {len(self.group_names)}"
        self.view_info_text.insert(1.0, info_text)
        self.view_info_text.config(state=tk.DISABLED)

    def show_week_view(self):
        """Показ расписания на неделю для всех групп со всеми парами"""
        for item in self.view_tree.get_children():
            self.view_tree.delete(item)

        # Настраиваем колонки - убираем колонку "Часы"
        columns = ['Группа'] + self.days_of_week + ['Итого часов']
        self.view_tree['columns'] = columns

        for col in columns:
            self.view_tree.heading(col, text=col)
            if col == 'Группа':
                self.view_tree.column(col, width=100)
            elif col == 'Итого часов':
                self.view_tree.column(col, width=100)
            else:
                self.view_tree.column(col, width=220)

        # Заполняем данными
        for group in sorted(self.group_names):
            row = [group]
            total_week_hours = 0

            for day in self.days_of_week:
                day_schedule = self.schedule_data.get(group, {}).get(day, {})
                day_text = ""
                day_hours = 0

                # Показываем ВСЕ пары (1 до max_para)
                for para in range(1, self.max_para + 1):
                    para_str = str(para)
                    lesson = day_schedule.get(para_str, {})
                    if lesson:
                        subject = lesson.get('предмет', '')
                        teacher = self.shorten_teacher_name(lesson.get('преподаватель', ''))
                        room = lesson.get('кабинет', '')
                        hours = lesson.get('часы', 2)
                        # Добавляем номер пары и информацию
                        day_text += f"{para}. {subject}\n{teacher}\n{room}\n\n"
                        day_hours += hours

                row.append(day_text.strip())
                total_week_hours += day_hours

            row.append(str(total_week_hours))
            self.view_tree.insert('', tk.END, values=row)

        # Обновляем информацию
        self.view_info_text.config(state=tk.NORMAL)
        self.view_info_text.delete(1.0, tk.END)
        info_text = f"📅 Расписание на неделю (все пары)\n"
        info_text += f"📚 Всего групп: {len(self.group_names)}"
        self.view_info_text.insert(1.0, info_text)
        self.view_info_text.config(state=tk.DISABLED)

    def export_view_to_excel(self):
        mode = self.view_mode.get()
        if mode == "day":
            self.export_day_view_excel()
        else:
            self.export_week_view_excel()

    def export_day_view_excel(self):
        """Экспорт дневного просмотра в Excel без столбца Часы"""
        day = self.view_day_var.get()
        if not day:
            return

        file_path = filedialog.asksaveasfilename(
            title=f"Сохранить расписание на {day}",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"Расписание_все_группы_{day}.xlsx"
        )

        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = f"Расписание {day}"

            # Стили
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font_white = Font(name='Arial', size=11, bold=True, color="FFFFFF")
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            # Заголовок
            ws.merge_cells('A1:' + get_column_letter(self.max_para + 2) + '1')
            title_cell = ws['A1']
            title_cell.value = f"РАСПИСАНИЕ ЗАНЯТИЙ НА {day.upper()}"
            title_cell.font = Font(name='Arial', size=16, bold=True)
            title_cell.alignment = center_alignment

            # Шапка - без столбца "Часы"
            headers = ['Группа'] + [f'{i} пара' for i in range(1, self.max_para + 1)] + ['Итого часов']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col)
                cell.value = header
                cell.font = header_font_white
                cell.fill = header_fill
                cell.alignment = center_alignment
                cell.border = border

            # Данные
            row_num = 3
            for group in sorted(self.group_names):
                day_schedule = self.schedule_data.get(group, {}).get(day, {})
                ws.cell(row=row_num, column=1, value=group).font = Font(bold=True)
                ws.cell(row=row_num, column=1).alignment = center_alignment
                ws.cell(row=row_num, column=1).border = border

                total_hours = 0
                for para in range(1, self.max_para + 1):
                    para_str = str(para)
                    lesson = day_schedule.get(para_str, {})
                    if lesson:
                        # Без часов в ячейке
                        text = f"{lesson.get('предмет', '')}\n{self.shorten_teacher_name(lesson.get('преподаватель', ''))}\n{lesson.get('кабинет', '')}"
                        hours = lesson.get('часы', 2)
                        total_hours += hours
                    else:
                        text = ""

                    cell = ws.cell(row=row_num, column=para + 1)
                    cell.value = text
                    cell.alignment = left_alignment
                    cell.border = border
                    cell.font = Font(size=9)

                ws.cell(row=row_num, column=self.max_para + 2, value=total_hours).font = Font(bold=True)
                ws.cell(row=row_num, column=self.max_para + 2).alignment = center_alignment
                ws.cell(row=row_num, column=self.max_para + 2).border = border
                row_num += 1

            # Настройка ширины колонок
            ws.column_dimensions['A'].width = 15
            for i in range(2, self.max_para + 2):
                ws.column_dimensions[get_column_letter(i)].width = 35
            ws.column_dimensions[get_column_letter(self.max_para + 2)].width = 15

            # Высота строк
            ws.row_dimensions[1].height = 30
            ws.row_dimensions[2].height = 25
            for row in range(3, row_num):
                ws.row_dimensions[row].height = 50

            wb.save(file_path)
            messagebox.showinfo("Успех", f"Расписание сохранено в файл:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {str(e)}")

    def export_week_view_excel(self):
        """Экспорт недельного просмотра в Excel со всеми парами"""
        file_path = filedialog.asksaveasfilename(
            title="Сохранить расписание на неделю",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="Расписание_все_группы_неделя.xlsx"
        )

        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Расписание неделя"

            # Стили
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font_white = Font(name='Arial', size=11, bold=True, color="FFFFFF")
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            # Заголовок
            ws.merge_cells('A1:' + get_column_letter(len(self.days_of_week) + 2) + '1')
            title_cell = ws['A1']
            title_cell.value = "РАСПИСАНИЕ ЗАНЯТИЙ НА НЕДЕЛЮ"
            title_cell.font = Font(name='Arial', size=16, bold=True)
            title_cell.alignment = center_alignment

            # Шапка
            headers = ['Группа'] + self.days_of_week + ['Итого часов']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col)
                cell.value = header
                cell.font = header_font_white
                cell.fill = header_fill
                cell.alignment = center_alignment
                cell.border = border

            # Данные
            row_num = 3
            for group in sorted(self.group_names):
                ws.cell(row=row_num, column=1, value=group).font = Font(bold=True)
                ws.cell(row=row_num, column=1).alignment = center_alignment
                ws.cell(row=row_num, column=1).border = border

                total_week_hours = 0
                for col, day in enumerate(self.days_of_week, 2):
                    day_schedule = self.schedule_data.get(group, {}).get(day, {})
                    day_text = ""
                    day_hours = 0

                    # Показываем ВСЕ пары (1 до max_para)
                    for para in range(1, self.max_para + 1):
                        para_str = str(para)
                        lesson = day_schedule.get(para_str, {})
                        if lesson:
                            subject = lesson.get('предмет', '')
                            teacher = self.shorten_teacher_name(lesson.get('преподаватель', ''))
                            room = lesson.get('кабинет', '')
                            hours = lesson.get('часы', 2)
                            day_text += f"{para}. {subject}\n{teacher}\n{room}\n\n"
                            day_hours += hours

                    cell = ws.cell(row=row_num, column=col)
                    cell.value = day_text.strip()
                    cell.alignment = left_alignment
                    cell.border = border
                    cell.font = Font(size=9)
                    total_week_hours += day_hours

                ws.cell(row=row_num, column=len(self.days_of_week) + 2, value=total_week_hours).font = Font(bold=True)
                ws.cell(row=row_num, column=len(self.days_of_week) + 2).alignment = center_alignment
                ws.cell(row=row_num, column=len(self.days_of_week) + 2).border = border
                row_num += 1

            # Настройка ширины колонок
            ws.column_dimensions['A'].width = 15
            for i in range(2, len(self.days_of_week) + 2):
                ws.column_dimensions[get_column_letter(i)].width = 40
            ws.column_dimensions[get_column_letter(len(self.days_of_week) + 2)].width = 15

            # Высота строк (увеличиваем, чтобы поместились все пары)
            ws.row_dimensions[1].height = 30
            ws.row_dimensions[2].height = 25
            for row in range(3, row_num):
                ws.row_dimensions[row].height = 100  # Увеличиваем высоту для всех пар

            wb.save(file_path)
            messagebox.showinfo("Успех", f"Расписание сохранено в файл:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {str(e)}")

    def export_view_to_pdf(self):
        mode = self.view_mode.get()
        if mode == "day":
            self.export_day_view_pdf()
        else:
            self.export_week_view_pdf()

    def export_day_view_pdf(self):
        """Экспорт дневного просмотра в PDF без столбца Часы"""
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Ошибка", "Для экспорта в PDF установите reportlab")
            return

        day = self.view_day_var.get()
        if not day:
            return

        file_path = filedialog.asksaveasfilename(
            title=f"Сохранить расписание на {day}",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=f"Расписание_все_группы_{day}.pdf"
        )

        if not file_path:
            return

        try:
            doc = SimpleDocTemplate(file_path, pagesize=A3,
                                    leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                                    topMargin=1.5 * cm, bottomMargin=1.5 * cm)
            story = []

            try:
                font_name = 'Arial'
                try:
                    pdfmetrics.getFont(font_name)
                except:
                    font_paths = [
                        "C:/Windows/Fonts/arial.ttf",
                        "C:/Windows/Fonts/ariali.ttf",
                        "C:/Windows/Fonts/times.ttf",
                        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                        "/System/Library/Fonts/Helvetica.ttc"
                    ]
                    for font_path in font_paths:
                        if os.path.exists(font_path):
                            pdfmetrics.registerFont(TTFont('Arial', font_path))
                            break
            except:
                pass

            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                'Title',
                parent=styles['Title'],
                fontSize=18,
                alignment=1,
                spaceAfter=20,
                fontName='Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'
            )
            story.append(Paragraph(f"РАСПИСАНИЕ ЗАНЯТИЙ НА {day.upper()}", title_style))
            story.append(Spacer(1, 0.3 * cm))

            date_style = ParagraphStyle(
                'Date',
                parent=styles['Normal'],
                fontSize=12,
                alignment=1,
                spaceAfter=15,
                fontName='Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'
            )
            story.append(Paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}", date_style))
            story.append(Spacer(1, 0.3 * cm))

            # Данные для таблицы - без столбца "Часы"
            data = [['Группа'] + [f'{i} пара' for i in range(1, self.max_para + 1)] + ['Итого']]

            for group in sorted(self.group_names):
                day_schedule = self.schedule_data.get(group, {}).get(day, {})
                row = [group]
                total_hours = 0

                for para in range(1, self.max_para + 1):
                    para_str = str(para)
                    lesson = day_schedule.get(para_str, {})
                    if lesson:
                        subject = lesson.get('предмет', '')
                        teacher = self.shorten_teacher_name(lesson.get('преподаватель', ''))
                        room = lesson.get('кабинет', '')
                        hours = lesson.get('часы', 2)
                        # Без часов в ячейке
                        text = f"{subject}\n{teacher}\n{room}"
                        total_hours += hours
                    else:
                        text = ""
                    row.append(text)

                row.append(str(total_hours))
                data.append(row)

            col_widths = [60] + [85] * self.max_para + [50]
            table = Table(data, colWidths=col_widths, repeatRows=1)

            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#366092')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#333333')),
                ('LINEABOVE', (0, 1), (-1, 1), 1, colors.HexColor('#366092')),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor('#1a3a5c')),
                ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 2), (-1, 2), colors.white),
                ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 4), (-1, 4), colors.white),
                ('BACKGROUND', (0, 5), (-1, 5), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 6), (-1, 6), colors.white),
                ('BACKGROUND', (0, 7), (-1, 7), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 8), (-1, 8), colors.white),
                ('BACKGROUND', (0, 9), (-1, 9), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 10), (-1, 10), colors.white),
                ('FONTNAME', (-1, 1), (-1, -1), 'Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'),
                ('FONTSIZE', (-1, 1), (-1, -1), 10),
                ('TEXTCOLOR', (-1, 1), (-1, -1), colors.HexColor('#1a3a5c')),
                ('BACKGROUND', (-1, 1), (-1, -1), colors.HexColor('#f0f0f0')),
            ])

            table.setStyle(style)
            story.append(table)

            footer_style = ParagraphStyle(
                'Footer',
                parent=styles['Normal'],
                fontSize=9,
                alignment=1,
                spaceTop=15,
                fontName='Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'
            )
            story.append(Paragraph(f"Сгенерировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}", footer_style))

            doc.build(story)
            messagebox.showinfo("Успех", f"PDF создан: {file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF: {str(e)}")

    def export_week_view_pdf(self):
        """Экспорт недельного просмотра в PDF со всеми парами"""
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Ошибка", "Для экспорта в PDF установите reportlab")
            return

        file_path = filedialog.asksaveasfilename(
            title="Сохранить расписание на неделю",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile="Расписание_все_группы_неделя.pdf"
        )

        if not file_path:
            return

        try:
            doc = SimpleDocTemplate(file_path, pagesize=landscape(A3),
                                    leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                                    topMargin=1.5 * cm, bottomMargin=1.5 * cm)
            story = []

            try:
                font_name = 'Arial'
                try:
                    pdfmetrics.getFont(font_name)
                except:
                    font_paths = [
                        "C:/Windows/Fonts/arial.ttf",
                        "C:/Windows/Fonts/ariali.ttf",
                        "C:/Windows/Fonts/times.ttf",
                        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                        "/System/Library/Fonts/Helvetica.ttc"
                    ]
                    for font_path in font_paths:
                        if os.path.exists(font_path):
                            pdfmetrics.registerFont(TTFont('Arial', font_path))
                            break
            except:
                pass

            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                'Title',
                parent=styles['Title'],
                fontSize=18,
                alignment=1,
                spaceAfter=20,
                fontName='Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'
            )
            story.append(Paragraph("РАСПИСАНИЕ ЗАНЯТИЙ НА НЕДЕЛЮ", title_style))
            story.append(Spacer(1, 0.3 * cm))

            date_style = ParagraphStyle(
                'Date',
                parent=styles['Normal'],
                fontSize=12,
                alignment=1,
                spaceAfter=15,
                fontName='Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'
            )
            story.append(Paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}", date_style))
            story.append(Spacer(1, 0.3 * cm))

            # Данные для таблицы
            data = [['Группа'] + self.days_of_week + ['Итого']]

            for group in sorted(self.group_names):
                row = [group]
                total_week_hours = 0

                for day in self.days_of_week:
                    day_schedule = self.schedule_data.get(group, {}).get(day, {})
                    day_text = ""
                    day_hours = 0

                    # Показываем ВСЕ пары (1 до max_para)
                    for para in range(1, self.max_para + 1):
                        para_str = str(para)
                        lesson = day_schedule.get(para_str, {})
                        if lesson:
                            subject = lesson.get('предмет', '')
                            teacher = self.shorten_teacher_name(lesson.get('преподаватель', ''))
                            room = lesson.get('кабинет', '')
                            hours = lesson.get('часы', 2)
                            day_text += f"{para}. {subject}\n{teacher}\n{room}\n\n"
                            day_hours += hours

                    row.append(day_text.strip())
                    total_week_hours += day_hours

                row.append(str(total_week_hours))
                data.append(row)

            col_widths = [55] + [85] * len(self.days_of_week) + [45]
            table = Table(data, colWidths=col_widths, repeatRows=1)

            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#366092')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('FONTSIZE', (0, 1), (-1, -1), 7),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#333333')),
                ('LINEABOVE', (0, 1), (-1, 1), 1, colors.HexColor('#366092')),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor('#1a3a5c')),
                ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 2), (-1, 2), colors.white),
                ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 4), (-1, 4), colors.white),
                ('BACKGROUND', (0, 5), (-1, 5), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 6), (-1, 6), colors.white),
                ('BACKGROUND', (0, 7), (-1, 7), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 8), (-1, 8), colors.white),
                ('BACKGROUND', (0, 9), (-1, 9), colors.HexColor('#e8f0fe')),
                ('BACKGROUND', (0, 10), (-1, 10), colors.white),
                ('FONTNAME', (-1, 1), (-1, -1), 'Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'),
                ('FONTSIZE', (-1, 1), (-1, -1), 10),
                ('TEXTCOLOR', (-1, 1), (-1, -1), colors.HexColor('#1a3a5c')),
                ('BACKGROUND', (-1, 1), (-1, -1), colors.HexColor('#f0f0f0')),
            ])

            table.setStyle(style)
            story.append(table)

            footer_style = ParagraphStyle(
                'Footer',
                parent=styles['Normal'],
                fontSize=9,
                alignment=1,
                spaceTop=15,
                fontName='Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'
            )
            story.append(Paragraph(f"Сгенерировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}", footer_style))

            doc.build(story)
            messagebox.showinfo("Успех", f"PDF создан: {file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF: {str(e)}")

    def show_all_groups_day(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Расписание на день (все группы)")
        dialog.geometry("1200x700")
        dialog.transient(self.root)

        control_frame = ttk.Frame(dialog)
        control_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(control_frame, text="Выберите день:").pack(side=tk.LEFT, padx=5)
        day_var = tk.StringVar(value=self.days_of_week[0])
        day_combo = ttk.Combobox(control_frame, textvariable=day_var,
                                 values=self.days_of_week, width=15, state='readonly')
        day_combo.pack(side=tk.LEFT, padx=5)

        def update_dialog_view():
            day = day_var.get()
            for item in tree.get_children():
                tree.delete(item)

            columns = ['Группа'] + [f'{i} пара' for i in range(1, self.max_para + 1)] + ['Итого часов']
            tree['columns'] = columns

            for col in columns:
                tree.heading(col, text=col)
                if col == 'Группа':
                    tree.column(col, width=100)
                elif col == 'Итого часов':
                    tree.column(col, width=100)
                else:
                    tree.column(col, width=180)

            for group in sorted(self.group_names):
                day_schedule = self.schedule_data.get(group, {}).get(day, {})
                row = [group]
                total_hours = 0

                for para in range(1, self.max_para + 1):
                    para_str = str(para)
                    lesson = day_schedule.get(para_str, {})
                    if lesson:
                        text = f"{lesson.get('предмет', '')}\n{self.shorten_teacher_name(lesson.get('преподаватель', ''))}\n{lesson.get('кабинет', '')}"
                        hours = lesson.get('часы', 2)
                        total_hours += hours
                    else:
                        text = ""
                    row.append(text)

                row.append(str(total_hours))
                tree.insert('', tk.END, values=row)

        ttk.Button(control_frame, text="Показать", command=update_dialog_view).pack(side=tk.LEFT, padx=10)
        ttk.Button(control_frame, text="Экспорт в Excel",
                   command=lambda: self.export_day_view_excel()).pack(side=tk.LEFT, padx=5)
        if REPORTLAB_AVAILABLE:
            ttk.Button(control_frame, text="Экспорт в PDF",
                       command=lambda: self.export_day_view_pdf()).pack(side=tk.LEFT, padx=5)

        table_frame = ttk.Frame(dialog)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tree = ttk.Treeview(table_frame, show='headings', height=20)
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Button(dialog, text="Закрыть", command=dialog.destroy).pack(pady=10)

        update_dialog_view()

    def show_all_groups_week(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Расписание на неделю (все группы)")
        dialog.geometry("1400x700")
        dialog.transient(self.root)

        control_frame = ttk.Frame(dialog)
        control_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(control_frame, text="Недельное расписание всех групп").pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Экспорт в Excel",
                   command=lambda: self.export_week_view_excel()).pack(side=tk.LEFT, padx=10)
        if REPORTLAB_AVAILABLE:
            ttk.Button(control_frame, text="Экспорт в PDF",
                       command=lambda: self.export_week_view_pdf()).pack(side=tk.LEFT, padx=5)

        table_frame = ttk.Frame(dialog)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tree = ttk.Treeview(table_frame, show='headings', height=20)
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        columns = ['Группа'] + self.days_of_week + ['Итого часов']
        tree['columns'] = columns

        for col in columns:
            tree.heading(col, text=col)
            if col == 'Группа':
                tree.column(col, width=100)
            elif col == 'Итого часов':
                tree.column(col, width=100)
            else:
                tree.column(col, width=200)

        for group in sorted(self.group_names):
            row = [group]
            total_week_hours = 0

            for day in self.days_of_week:
                day_schedule = self.schedule_data.get(group, {}).get(day, {})
                day_text = ""
                day_hours = 0

                for para in range(1, self.max_para + 1):
                    para_str = str(para)
                    lesson = day_schedule.get(para_str, {})
                    if lesson:
                        day_text += f"{para}. {lesson.get('предмет', '')[:15]}\n"
                        day_text += f"{self.shorten_teacher_name(lesson.get('преподаватель', ''))}\n"
                        day_text += f"{lesson.get('кабинет', '')}\n\n"
                        day_hours += lesson.get('часы', 2)

                row.append(day_text.strip())
                total_week_hours += day_hours

            row.append(str(total_week_hours))
            tree.insert('', tk.END, values=row)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Button(dialog, text="Закрыть", command=dialog.destroy).pack(pady=10)

    def export_week_schedule(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Экспорт расписания на неделю")
        dialog.geometry("500x300")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Выберите формат экспорта:", font=('Arial', 12)).pack(pady=20)

        format_var = tk.StringVar(value="excel")
        ttk.Radiobutton(dialog, text="Excel", variable=format_var, value="excel",
                        font=('Arial', 11)).pack(pady=10)
        if REPORTLAB_AVAILABLE:
            ttk.Radiobutton(dialog, text="PDF", variable=format_var, value="pdf",
                            font=('Arial', 11)).pack(pady=10)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)

        def export_week():
            file_format = format_var.get()
            dialog.destroy()

            if file_format == 'excel':
                self.export_week_view_excel()
            else:
                self.export_week_view_pdf()

        ttk.Button(btn_frame, text="Экспортировать", command=export_week, width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Отмена", command=dialog.destroy, width=20).pack(side=tk.LEFT, padx=10)

    def create_teachers_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=5)

        ttk.Label(control_frame, text="Преподаватель:").pack(side=tk.LEFT, padx=5)

        self.teacher_var = tk.StringVar()
        self.teacher_combo = ttk.Combobox(control_frame, textvariable=self.teacher_var,
                                          values=[], width=25, state='readonly')
        self.teacher_combo.pack(side=tk.LEFT, padx=5)
        self.teacher_combo.bind('<<ComboboxSelected>>', self.on_teacher_change)

        ttk.Button(control_frame, text="Обновить", command=self.update_teachers_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Экспорт", command=self.export_teachers_hours).pack(side=tk.LEFT, padx=5)

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        columns = ('День', 'Пара', 'Группа', 'Предмет', 'Кабинет', 'Часы')
        self.teachers_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)

        self.teachers_tree.heading('День', text='День')
        self.teachers_tree.heading('Пара', text='Пара')
        self.teachers_tree.heading('Группа', text='Группа')
        self.teachers_tree.heading('Предмет', text='Предмет')
        self.teachers_tree.heading('Кабинет', text='Кабинет')
        self.teachers_tree.heading('Часы', text='Часы')

        self.teachers_tree.column('День', width=120)
        self.teachers_tree.column('Пара', width=60)
        self.teachers_tree.column('Группа', width=100)
        self.teachers_tree.column('Предмет', width=350)
        self.teachers_tree.column('Кабинет', width=120)
        self.teachers_tree.column('Часы', width=60)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.teachers_tree.yview)
        self.teachers_tree.configure(yscrollcommand=scrollbar.set)

        self.teachers_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        info_frame = ttk.LabelFrame(main_frame, text="Информация")
        info_frame.pack(fill=tk.X, pady=5)

        self.teacher_info_text = tk.Text(info_frame, height=4, width=80, state=tk.DISABLED, font=('Arial', 9))
        self.teacher_info_text.pack(fill=tk.X, padx=5, pady=5)

        self.update_teachers_list()

    def create_rooms_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=5)

        ttk.Label(control_frame, text="Кабинет:").pack(side=tk.LEFT, padx=5)

        self.room_var = tk.StringVar()
        self.room_combo = ttk.Combobox(control_frame, textvariable=self.room_var,
                                       values=[], width=15, state='readonly')
        self.room_combo.pack(side=tk.LEFT, padx=5)
        self.room_combo.bind('<<ComboboxSelected>>', self.on_room_change)

        ttk.Button(control_frame, text="Обновить", command=self.update_rooms_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Статистика", command=self.show_rooms_stats).pack(side=tk.LEFT, padx=5)

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        columns = ('День', 'Пара', 'Группа', 'Предмет', 'Преподаватель', 'Часы')
        self.rooms_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)

        self.rooms_tree.heading('День', text='День')
        self.rooms_tree.heading('Пара', text='Пара')
        self.rooms_tree.heading('Группа', text='Группа')
        self.rooms_tree.heading('Предмет', text='Предмет')
        self.rooms_tree.heading('Преподаватель', text='Преподаватель')
        self.rooms_tree.heading('Часы', text='Часы')

        self.rooms_tree.column('День', width=120)
        self.rooms_tree.column('Пара', width=60)
        self.rooms_tree.column('Группа', width=100)
        self.rooms_tree.column('Предмет', width=300)
        self.rooms_tree.column('Преподаватель', width=250)
        self.rooms_tree.column('Часы', width=60)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.rooms_tree.yview)
        self.rooms_tree.configure(yscrollcommand=scrollbar.set)

        self.rooms_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        info_frame = ttk.LabelFrame(main_frame, text="Информация")
        info_frame.pack(fill=tk.X, pady=5)

        self.room_info_text = tk.Text(info_frame, height=4, width=80, state=tk.DISABLED, font=('Arial', 9))
        self.room_info_text.pack(fill=tk.X, padx=5, pady=5)

        self.update_rooms_list()

    def create_stats_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=5)

        ttk.Button(control_frame, text="Обновить", command=self.update_stats).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Экспорт в Excel", command=self.export_stats_to_excel).pack(side=tk.LEFT, padx=5)

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        columns = ('Показатель', 'Значение')
        self.stats_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)

        self.stats_tree.heading('Показатель', text='Показатель')
        self.stats_tree.heading('Значение', text='Значение')

        self.stats_tree.column('Показатель', width=300)
        self.stats_tree.column('Значение', width=200)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.stats_tree.yview)
        self.stats_tree.configure(yscrollcommand=scrollbar.set)

        self.stats_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        if MATPLOTLIB_AVAILABLE:
            chart_frame = ttk.LabelFrame(main_frame, text="Графики")
            chart_frame.pack(fill=tk.BOTH, expand=True, pady=5)

            self.figure = Figure(figsize=(8, 4), dpi=100)
            self.canvas = FigureCanvasTkAgg(self.figure, master=chart_frame)
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.update_stats()

    def create_settings_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        general_frame = ttk.LabelFrame(main_frame, text="Общие настройки")
        general_frame.pack(fill=tk.X, pady=5)

        ttk.Label(general_frame, text="Максимальное количество пар:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.max_para_var = tk.StringVar(value=str(self.max_para))
        max_para_spin = ttk.Spinbox(general_frame, from_=4, to=10, textvariable=self.max_para_var, width=10)
        max_para_spin.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(general_frame, text="Применить", command=self.apply_max_para).grid(row=0, column=2, padx=5, pady=5)

        backup_frame = ttk.LabelFrame(main_frame, text="Резервное копирование")
        backup_frame.pack(fill=tk.X, pady=5)

        ttk.Label(backup_frame, text="Папка с бэкапами:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(backup_frame, text=self.backup_folder).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Button(backup_frame, text="Создать бэкап", command=self.create_backup_manual).grid(row=0, column=2, padx=5,
                                                                                               pady=5)
        ttk.Button(backup_frame, text="Восстановить", command=self.restore_backup).grid(row=0, column=3, padx=5, pady=5)

        info_frame = ttk.LabelFrame(main_frame, text="Информация")
        info_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        info_text = tk.Text(info_frame, height=10, width=80, state=tk.DISABLED, font=('Arial', 9))
        info_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        info_text.config(state=tk.NORMAL)
        info_text.insert(1.0,
                         "Горячие клавиши:\n"
                         "  Ctrl+N - Добавить пару\n"
                         "  Ctrl+S - Сохранить\n"
                         "  Ctrl+F - Поиск\n"
                         "  Delete - Удалить пару\n"
                         "  Ctrl+Z - Отменить\n\n"
                         "Форматы файлов:\n"
                         "  Импорт: .xlsx, .xls\n"
                         "  Экспорт: .xlsx, .pdf (при установленном reportlab)\n\n"
                         "Для экспорта в PDF установите:\n"
                         "  pip install reportlab\n"
                         "Для графиков установите:\n"
                         "  pip install matplotlib")
        info_text.config(state=tk.DISABLED)

    def apply_max_para(self):
        try:
            new_max = int(self.max_para_var.get())
            if 4 <= new_max <= 10:
                self.max_para = new_max
                self.update_editor_table()
                messagebox.showinfo("Успех", f"Максимальное количество пар установлено: {new_max}")
            else:
                messagebox.showwarning("Внимание", "Количество пар должно быть от 4 до 10")
        except:
            messagebox.showerror("Ошибка", "Введите корректное число")

    def update_stats(self):
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)

        total_groups = len(self.group_names)
        total_lessons = 0
        total_hours = 0
        teacher_count = set()
        room_count = set()

        for group in self.schedule_data:
            for day in self.days_of_week:
                for lesson in self.schedule_data.get(group, {}).get(day, {}).values():
                    total_lessons += 1
                    total_hours += lesson.get('часы', 2)
                    teacher = lesson.get('преподаватель', '')
                    if teacher:
                        for t in re.split(r'[/,]', teacher):
                            t = t.strip()
                            if t:
                                teacher_count.add(t)
                    room = lesson.get('кабинет', '')
                    if room:
                        room_count.add(room)

        stats = [
            ("Всего групп", str(total_groups)),
            ("Всего занятий", str(total_lessons)),
            ("Всего часов", str(total_hours)),
            ("Среднее часов на группу", f"{total_hours / total_groups:.1f}" if total_groups > 0 else "0"),
            ("Количество преподавателей", str(len(teacher_count))),
            ("Количество кабинетов", str(len(room_count))),
            ("Среднее занятий на группу", f"{total_lessons / total_groups:.1f}" if total_groups > 0 else "0"),
        ]

        for stat, value in stats:
            self.stats_tree.insert('', tk.END, values=(stat, value))

        if MATPLOTLIB_AVAILABLE:
            self.create_charts()

    def create_charts(self):
        self.figure.clear()

        ax1 = self.figure.add_subplot(121)
        groups_hours = {}
        for group in self.group_names:
            hours = 0
            for day in self.days_of_week:
                for lesson in self.schedule_data.get(group, {}).get(day, {}).values():
                    hours += lesson.get('часы', 2)
            groups_hours[group] = hours

        if groups_hours:
            groups = list(groups_hours.keys())
            hours = list(groups_hours.values())
            ax1.bar(groups, hours, color='skyblue')
            ax1.set_xlabel('Группы')
            ax1.set_ylabel('Часы')
            ax1.set_title('Нагрузка по группам')
            ax1.tick_params(axis='x', rotation=45)

        ax2 = self.figure.add_subplot(122)
        days_hours = {day: 0 for day in self.days_of_week}
        for group in self.schedule_data:
            for day in self.days_of_week:
                for lesson in self.schedule_data.get(group, {}).get(day, {}).values():
                    days_hours[day] += lesson.get('часы', 2)

        if any(days_hours.values()):
            days = list(days_hours.keys())
            hours = list(days_hours.values())
            ax2.bar(days, hours, color='lightgreen')
            ax2.set_xlabel('Дни')
            ax2.set_ylabel('Часы')
            ax2.set_title('Нагрузка по дням')
            ax2.tick_params(axis='x', rotation=45)

        self.figure.tight_layout()
        self.canvas.draw()

    def export_stats_to_excel(self):
        file_path = filedialog.asksaveasfilename(
            title="Сохранить статистику",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            rows = []
            for item in self.stats_tree.get_children():
                values = self.stats_tree.item(item, 'values')
                rows.append({
                    'Показатель': values[0],
                    'Значение': values[1]
                })

            if rows:
                df = pd.DataFrame(rows)
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Успех", f"Статистика экспортирована в {file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать статистику: {str(e)}")

    def update_rooms_list(self):
        rooms = set()
        for group in self.schedule_data:
            for day in self.schedule_data[group]:
                for lesson in self.schedule_data[group][day].values():
                    room = lesson.get('кабинет', '')
                    if room and room.strip():
                        rooms.add(room)

        self.room_combo['values'] = sorted(rooms)
        if rooms and not self.room_var.get():
            self.room_var.set(sorted(rooms)[0])
            self.show_room_schedule(sorted(rooms)[0])

    def on_room_change(self, event):
        room = self.room_var.get()
        if room:
            self.show_room_schedule(room)

    def show_room_schedule(self, room_name):
        for item in self.rooms_tree.get_children():
            self.rooms_tree.delete(item)

        lessons = []
        total_hours = 0

        for group in self.schedule_data:
            for day in self.days_of_week:
                for para, lesson in self.schedule_data.get(group, {}).get(day, {}).items():
                    if lesson.get('кабинет', '') == room_name:
                        hours = lesson.get('часы', 2)
                        lessons.append({
                            'день': day,
                            'пара': para,
                            'группа': group,
                            'предмет': lesson.get('предмет', ''),
                            'преподаватель': lesson.get('преподаватель', ''),
                            'часы': hours
                        })
                        total_hours += hours

        day_order = {day: i for i, day in enumerate(self.days_of_week)}
        lessons.sort(key=lambda x: (day_order.get(x['день'], 99), int(x['пара']) if x['пара'].isdigit() else 99))

        for lesson in lessons:
            self.rooms_tree.insert('', tk.END, values=(
                lesson['день'],
                lesson['пара'],
                lesson['группа'],
                lesson['предмет'],
                lesson['преподаватель'],
                lesson['часы']
            ))

        self.room_info_text.config(state=tk.NORMAL)
        self.room_info_text.delete(1.0, tk.END)

        info_text = f"🏢 Кабинет: {room_name}\n"
        info_text += f"📚 Всего занятий: {len(lessons)}\n"
        info_text += f"⏱ Всего часов: {total_hours}\n"

        self.room_info_text.insert(1.0, info_text)
        self.room_info_text.config(state=tk.DISABLED)

    def show_rooms_stats(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Статистика по кабинетам")
        dialog.geometry("650x450")
        dialog.transient(self.root)

        text = ScrolledText(dialog, width=80, height=20, font=('Arial', 10))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        text.insert(tk.END, "📊 СТАТИСТИКА ПО КАБИНЕТАМ\n", 'header')
        text.insert(tk.END, "=" * 60 + "\n\n", 'header')

        room_stats = defaultdict(lambda: {'lessons': 0, 'hours': 0, 'groups': set()})

        for group in self.schedule_data:
            for day in self.days_of_week:
                for lesson in self.schedule_data.get(group, {}).get(day, {}).values():
                    room = lesson.get('кабинет', '')
                    if room:
                        room_stats[room]['lessons'] += 1
                        room_stats[room]['hours'] += lesson.get('часы', 2)
                        room_stats[room]['groups'].add(group)

        for room, stats in sorted(room_stats.items(), key=lambda x: x[1]['hours'], reverse=True):
            text.insert(tk.END,
                        f"🏢 {room}\n"
                        f"  Занятий: {stats['lessons']}\n"
                        f"  Часов: {stats['hours']}\n"
                        f"  Групп: {', '.join(sorted(stats['groups']))}\n\n")

        text.tag_config('header', foreground='blue', font=('Arial', 12, 'bold'))
        text.config(state=tk.DISABLED)

        ttk.Button(dialog, text="Закрыть", command=dialog.destroy).pack(pady=10)

    def export_to_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Ошибка", "Для экспорта в PDF установите reportlab")
            return

        try:
            file_path = filedialog.asksaveasfilename(
                title="Сохранить PDF",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )

            if not file_path:
                return

            doc = SimpleDocTemplate(file_path, pagesize=A3)
            story = []
            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                'Title',
                parent=styles['Title'],
                fontSize=18,
                alignment=1,
                spaceAfter=20,
                fontName='Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'
            )
            story.append(Paragraph("РАСПИСАНИЕ ЗАНЯТИЙ", title_style))

            data = [['Группа'] + self.days_of_week]

            for group in sorted(self.group_names):
                row = [group]
                for day in self.days_of_week:
                    day_schedule = self.schedule_data.get(group, {}).get(day, {})
                    text = ""
                    for para, lesson in sorted(day_schedule.items()):
                        if lesson:
                            subject = lesson.get('предмет', '')[:25]
                            teacher = self.shorten_teacher_name(lesson.get('преподаватель', ''))
                            room = lesson.get('кабинет', '')
                            hours = lesson.get('часы', 2)
                            text += f"{para}. {subject}\n{teacher}\n{room} ({hours} ч.)\n\n"
                    row.append(text.strip())
                data.append(row)

            col_widths = [60] + [85] * 6
            table = Table(data, colWidths=col_widths)

            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ])

            table.setStyle(style)
            story.append(table)

            doc.build(story)
            messagebox.showinfo("Успех", f"PDF создан: {file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF: {str(e)}")

    def export_to_google_calendar(self):
        try:
            file_path = filedialog.asksaveasfilename(
                title="Сохранить для Google Calendar",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )

            if not file_path:
                return

            rows = []
            for group in self.schedule_data:
                for day in self.days_of_week:
                    for para, lesson in self.schedule_data.get(group, {}).get(day, {}).items():
                        if lesson:
                            subject = lesson.get('предмет', '')
                            teacher = lesson.get('преподаватель', '')
                            room = lesson.get('кабинет', '')
                            hours = lesson.get('часы', 2)

                            para_num = int(para) if para.isdigit() else 1
                            start_hour = 8 + (para_num - 1) * 2
                            end_hour = start_hour + hours

                            rows.append({
                                'Subject': f"{subject} ({group})",
                                'Start Date': self.get_next_date(day),
                                'Start Time': f"{start_hour:02d}:00",
                                'End Time': f"{end_hour:02d}:00",
                                'Location': room,
                                'Description': f"Преподаватель: {teacher}\nГруппа: {group}"
                            })

            if rows:
                df = pd.DataFrame(rows)
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
                messagebox.showinfo("Успех",
                                    f"Файл для Google Calendar создан: {file_path}\n\n"
                                    "Инструкция по импорту:\n"
                                    "1. Откройте Google Calendar\n"
                                    "2. Нажмите на '+' рядом с 'Другие календари'\n"
                                    "3. Выберите 'Импорт'\n"
                                    "4. Выберите созданный CSV файл")

                try:
                    os.startfile(os.path.dirname(file_path))
                except:
                    pass
            else:
                messagebox.showwarning("Внимание", "Нет данных для экспорта")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать: {str(e)}")

    def get_next_date(self, day_name):
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']
        today = datetime.now()
        today_day = today.weekday()

        day_map = {d: i for i, d in enumerate(days)}
        target_day = day_map.get(day_name, 0)

        days_until = target_day - today_day
        if days_until <= 0:
            days_until += 7

        target_date = today + timedelta(days=days_until)
        return target_date.strftime("%m/%d/%Y")

    def refresh_all(self):
        self.update_schedule_display()
        self.update_editor_table()
        self.update_teachers_list()
        self.update_rooms_list()
        self.update_stats()
        self.update_history_display()

    def create_backup_manual(self):
        self.create_backup()
        messagebox.showinfo("Успех", "Бэкап создан в папке 'backups'")

    def show_shortcuts(self):
        messagebox.showinfo("Горячие клавиши",
                            "Ctrl+N - Добавить пару\n"
                            "Ctrl+S - Сохранить расписание\n"
                            "Ctrl+F - Поиск\n"
                            "Delete - Удалить выбранную пару\n"
                            "Ctrl+Z - Отменить действие")

    def show_about(self):
        messagebox.showinfo("О программе",
                            "Управление расписанием\n"
                            "Версия 2.1\n\n"
                            "Программа для составления и управления расписанием\n"
                            "в учебных заведениях.\n\n"
                            "Особенности:\n"
                            "• Импорт/экспорт данных\n"
                            "• Редактирование расписания\n"
                            "• Подсчет нагрузки преподавателей\n"
                            "• Визуальная статистика\n"
                            "• Резервное копирование\n"
                            "• Экспорт в PDF с поддержкой кириллицы\n\n"
                            "© 2024 Все права защищены")

    def export_day_schedule(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Выбор дня для вывода расписания")
        dialog.geometry("500x500")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Выберите день недели:", font=('Arial', 14, 'bold')).pack(pady=15)

        days_frame = ttk.LabelFrame(dialog, text="Дни недели")
        days_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        all_days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']

        day_var = tk.StringVar()
        day_var.set('Понедельник')

        for i, day in enumerate(all_days):
            row = i // 2
            col = i % 2
            ttk.Radiobutton(days_frame, text=day, variable=day_var, value=day,
                            font=('Arial', 11)).grid(row=row, column=col, padx=20, pady=8, sticky=tk.W)

        format_frame = ttk.LabelFrame(dialog, text="Формат вывода")
        format_frame.pack(fill=tk.X, padx=20, pady=10)

        format_var = tk.StringVar(value="excel")
        format_inner = ttk.Frame(format_frame)
        format_inner.pack(pady=10)

        ttk.Radiobutton(format_inner, text="Excel", variable=format_var, value="excel",
                        font=('Arial', 11)).pack(side=tk.LEFT, padx=20)
        if REPORTLAB_AVAILABLE:
            ttk.Radiobutton(format_inner, text="PDF", variable=format_var, value="pdf",
                            font=('Arial', 11)).pack(side=tk.LEFT, padx=20)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)

        def export_schedule():
            selected_day = day_var.get()
            file_format = format_var.get()
            dialog.destroy()

            if selected_day == 'Воскресенье':
                if messagebox.askyesno("Воскресенье",
                                       "В воскресенье занятий обычно нет.\n"
                                       "Продолжить создание пустого расписания?"):
                    self.generate_day_schedule_excel(selected_day)
                return

            if file_format == 'excel':
                self.generate_day_schedule_excel(selected_day)
            else:
                self.generate_day_schedule_pdf(selected_day)

        ttk.Button(btn_frame, text="Экспортировать", command=export_schedule,
                   width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Отмена", command=dialog.destroy,
                   width=20).pack(side=tk.LEFT, padx=10)

    def generate_day_schedule_excel(self, day):
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = f"Расписание {day}"

            header_font = Font(name='Arial', size=12, bold=True)
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font_white = Font(name='Arial', size=12, bold=True, color="FFFFFF")
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            ws.merge_cells('A1:F1')
            title_cell = ws['A1']
            title_cell.value = f"РАСПИСАНИЕ ЗАНЯТИЙ НА {day.upper()}"
            title_cell.font = Font(name='Arial', size=16, bold=True)
            title_cell.alignment = center_alignment

            headers = ['Группа', '1 пара', '2 пара', '3 пара', '4 пара', '5 пара']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col)
                cell.value = header
                cell.font = header_font_white
                cell.fill = header_fill
                cell.alignment = center_alignment
                cell.border = border

            groups = sorted(self.group_names)
            current_row = 3

            for group in groups:
                day_schedule = self.schedule_data.get(group, {}).get(day, {})

                ws.cell(row=current_row, column=1, value=group).font = Font(bold=True)
                ws.cell(row=current_row, column=1).alignment = center_alignment
                ws.cell(row=current_row, column=1).border = border

                for para in range(1, 6):
                    para_str = str(para)
                    lesson = day_schedule.get(para_str, {})

                    if lesson:
                        subject = lesson.get('предмет', '')
                        teacher = lesson.get('преподаватель', '')
                        room = lesson.get('кабинет', '')
                        hours = lesson.get('часы', 2)

                        teacher_short = self.shorten_teacher_name(teacher)
                        text = f"{subject}\n{teacher_short}\n{room} ({hours} ч.)"
                    else:
                        text = ""

                    cell = ws.cell(row=current_row, column=para + 1)
                    cell.value = text
                    cell.alignment = left_alignment
                    cell.border = border
                    cell.font = Font(size=10)

                current_row += 1

            ws.column_dimensions['A'].width = 15
            for col in ['B', 'C', 'D', 'E', 'F']:
                ws.column_dimensions[col].width = 35

            ws.row_dimensions[1].height = 30
            ws.row_dimensions[2].height = 25
            for row in range(3, current_row):
                ws.row_dimensions[row].height = 60

            file_path = filedialog.asksaveasfilename(
                title=f"Сохранить расписание на {day}",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=f"Расписание_{day}.xlsx"
            )

            if file_path:
                wb.save(file_path)
                messagebox.showinfo("Успех", f"Расписание на {day} сохранено в файл:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать расписание: {str(e)}")

    def generate_day_schedule_pdf(self, day):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Ошибка", "Для экспорта в PDF установите reportlab")
            return

        try:
            file_path = filedialog.asksaveasfilename(
                title=f"Сохранить расписание на {day}",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                initialfile=f"Расписание_{day}.pdf"
            )

            if not file_path:
                return

            doc = SimpleDocTemplate(file_path, pagesize=A3,
                                    leftMargin=1 * cm, rightMargin=1 * cm,
                                    topMargin=1 * cm, bottomMargin=1 * cm)
            story = []

            styles = getSampleStyleSheet()

            try:
                font_name = 'Arial'
                try:
                    pdfmetrics.getFont(font_name)
                except:
                    font_paths = [
                        "C:/Windows/Fonts/arial.ttf",
                        "C:/Windows/Fonts/ariali.ttf",
                        "C:/Windows/Fonts/times.ttf",
                        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                        "/System/Library/Fonts/Helvetica.ttc"
                    ]
                    for font_path in font_paths:
                        if os.path.exists(font_path):
                            pdfmetrics.registerFont(TTFont('Arial', font_path))
                            break
            except:
                pass

            title_style = ParagraphStyle(
                'Title',
                parent=styles['Title'],
                fontSize=16,
                alignment=1,
                spaceAfter=20,
                fontName='Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'
            )

            story.append(Paragraph(f"РАСПИСАНИЕ ЗАНЯТИЙ НА {day.upper()}", title_style))
            story.append(Spacer(1, 0.5 * cm))

            data = [['Группа', '1 пара', '2 пара', '3 пара', '4 пара', '5 пара']]

            for group in sorted(self.group_names):
                row = [group]
                day_schedule = self.schedule_data.get(group, {}).get(day, {})
                for para in range(1, 6):
                    para_str = str(para)
                    lesson = day_schedule.get(para_str, {})
                    if lesson:
                        subject = lesson.get('предмет', '')
                        teacher = self.shorten_teacher_name(lesson.get('преподаватель', ''))
                        room = lesson.get('кабинет', '')
                        hours = lesson.get('часы', 2)
                        text = f"{subject}\n{teacher}\n{room} ({hours} ч.)"
                        row.append(text)
                    else:
                        row.append("")
                data.append(row)

            table = Table(data, colWidths=[60, 85, 85, 85, 85, 85])

            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial' if 'Arial' in pdfmetrics._fonts else 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ])

            table.setStyle(style)
            story.append(table)

            doc.build(story)
            messagebox.showinfo("Успех", f"PDF создан: {file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF: {str(e)}")

    def shorten_teacher_name(self, full_name):
        if not full_name:
            return ""

        parts = full_name.split()
        if len(parts) >= 3:
            return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
        elif len(parts) >= 2:
            return f"{parts[0]} {parts[1][0]}."
        else:
            return full_name

    def on_editor_group_change(self, event):
        self.update_editor_table()

    def update_editor_table(self):
        for item in self.editor_tree.get_children():
            self.editor_tree.delete(item)

        group = self.editor_group_var.get()
        if not group or group not in self.schedule_data:
            return

        for para_num in range(1, self.max_para + 1):
            para_str = str(para_num)
            values = [para_str]

            for day in self.days_of_week:
                lesson = self.schedule_data.get(group, {}).get(day, {}).get(para_str, {})
                if lesson:
                    text = f"{lesson.get('предмет', '')}\n{lesson.get('преподаватель', '')}\n{lesson.get('кабинет', '')} ({lesson.get('часы', 2)} ч.)"
                else:
                    text = ""
                values.append(text)

            self.editor_tree.insert('', tk.END, values=values)

    def edit_cell(self, event):
        item = self.editor_tree.selection()[0] if self.editor_tree.selection() else None
        if not item:
            return

        column = self.editor_tree.identify_column(event.x)
        col_index = int(column.replace('#', '')) - 1

        if col_index == 0:
            return

        day = self.days_of_week[col_index - 1]
        para = self.editor_tree.item(item, 'values')[0]
        group = self.editor_group_var.get()

        if not group or group not in self.schedule_data:
            return

        lesson = self.schedule_data.get(group, {}).get(day, {}).get(para, {})

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Редактирование - {group}, {day}, пара {para}")
        dialog.geometry("550x400")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Предмет:", font=('Arial', 10)).pack(pady=5)
        subject_var = tk.StringVar(value=lesson.get('предмет', ''))
        subject_combo = ttk.Combobox(dialog, textvariable=subject_var,
                                     values=self.subjects_list, width=45)
        subject_combo.pack(pady=5)

        ttk.Label(dialog, text="Преподаватель:", font=('Arial', 10)).pack(pady=5)
        teacher_var = tk.StringVar(value=lesson.get('преподаватель', ''))
        teacher_combo = ttk.Combobox(dialog, textvariable=teacher_var,
                                     values=self.teachers_list, width=45)
        teacher_combo.pack(pady=5)

        ttk.Label(dialog, text="Кабинет:", font=('Arial', 10)).pack(pady=5)
        room_var = tk.StringVar(value=lesson.get('кабинет', ''))
        room_entry = ttk.Entry(dialog, textvariable=room_var, width=45)
        room_entry.pack(pady=5)

        ttk.Label(dialog, text="Часы (1 пара = 2 часа):", font=('Arial', 10)).pack(pady=5)
        hours_var = tk.StringVar(value=str(lesson.get('часы', 2)))
        hours_entry = ttk.Entry(dialog, textvariable=hours_var, width=10)
        hours_entry.pack(pady=5)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)

        def save_cell():
            subject = subject_var.get().strip()
            teacher = teacher_var.get().strip()
            room = room_var.get().strip()
            hours = float(hours_var.get()) if hours_var.get() else 2

            if subject:
                if group not in self.schedule_data:
                    self.schedule_data[group] = {d: {} for d in self.days_of_week}
                if day not in self.schedule_data[group]:
                    self.schedule_data[group][day] = {}

                self.schedule_data[group][day][para] = {
                    'предмет': subject,
                    'преподаватель': teacher,
                    'кабинет': room,
                    'часы': hours
                }
            else:
                if group in self.schedule_data and day in self.schedule_data[group]:
                    if para in self.schedule_data[group][day]:
                        del self.schedule_data[group][day][para]

            dialog.destroy()
            self.update_editor_table()
            self.update_schedule_display()
            self.save_to_excel()

        ttk.Button(btn_frame, text="💾 Сохранить", command=save_cell, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑 Очистить",
                   command=lambda: [subject_var.set(""), teacher_var.set(""), room_var.set(""), hours_var.set("2")],
                   width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ Отмена", command=dialog.destroy, width=15).pack(side=tk.LEFT, padx=5)

    def create_weekly_schedule(self):
        group = self.editor_group_var.get()
        if not group:
            messagebox.showwarning("Внимание", "Выберите группу")
            return

        if group not in self.schedule_data:
            self.schedule_data[group] = {day: {} for day in self.days_of_week}

        self.update_editor_table()
        messagebox.showinfo("Успех", f"Создано расписание для группы {group}")

    def clear_weekly_schedule(self):
        group = self.editor_group_var.get()
        if not group:
            return

        if messagebox.askyesno("Подтверждение", f"Очистить все расписание для группы {group}?"):
            if group in self.schedule_data:
                self.schedule_data[group] = {day: {} for day in self.days_of_week}
            self.update_editor_table()
            self.update_schedule_display()
            self.save_to_excel()
            messagebox.showinfo("Успех", f"Расписание для группы {group} очищено")

    def save_editor_schedule(self):
        self.update_schedule_display()
        self.save_to_excel()
        self.save_current_schedule()
        messagebox.showinfo("Успех", "Расписание сохранено")

    def update_teachers_list(self):
        teachers = set()
        for group in self.schedule_data:
            for day in self.schedule_data[group]:
                for lesson in self.schedule_data[group][day].values():
                    teacher = lesson.get('преподаватель', '')
                    if teacher and teacher.strip():
                        for t in re.split(r'[/,]', teacher):
                            t = t.strip()
                            if t and t not in teachers:
                                teachers.add(t)

        self.teacher_combo['values'] = sorted(teachers)
        if teachers and not self.teacher_var.get():
            self.teacher_var.set(sorted(teachers)[0])
            self.show_teacher_schedule(sorted(teachers)[0])

    def on_teacher_change(self, event):
        teacher = self.teacher_var.get()
        if teacher:
            self.show_teacher_schedule(teacher)

    def show_teacher_schedule(self, teacher_name):
        for item in self.teachers_tree.get_children():
            self.teachers_tree.delete(item)

        lessons = []
        total_hours = 0
        days_count = defaultdict(int)

        for group in self.schedule_data:
            for day in self.days_of_week:
                for para, lesson in self.schedule_data.get(group, {}).get(day, {}).items():
                    teachers = re.split(r'[/,]', lesson.get('преподаватель', ''))
                    for t in teachers:
                        if t.strip() == teacher_name:
                            hours = lesson.get('часы', 2)
                            lessons.append({
                                'день': day,
                                'пара': para,
                                'группа': group,
                                'предмет': lesson.get('предмет', ''),
                                'кабинет': lesson.get('кабинет', ''),
                                'часы': hours
                            })
                            total_hours += hours
                            days_count[day] += hours

        day_order = {day: i for i, day in enumerate(self.days_of_week)}
        lessons.sort(key=lambda x: (day_order.get(x['день'], 99), int(x['пара']) if x['пара'].isdigit() else 99))

        for lesson in lessons:
            self.teachers_tree.insert('', tk.END, values=(
                lesson['день'],
                lesson['пара'],
                lesson['группа'],
                lesson['предмет'],
                lesson['кабинет'],
                lesson['часы']
            ))

        self.teacher_info_text.config(state=tk.NORMAL)
        self.teacher_info_text.delete(1.0, tk.END)

        info_text = f"👨‍🏫 Преподаватель: {teacher_name}\n"
        info_text += f"📚 Всего занятий: {len(lessons)}\n"
        info_text += f"⏱ Всего часов за неделю: {total_hours}\n"

        if days_count:
            info_text += "\n📅 Распределение по дням:\n"
            for day in self.days_of_week:
                if day in days_count:
                    info_text += f"  {day}: {days_count[day]} часов\n"

        self.teacher_info_text.insert(1.0, info_text)
        self.teacher_info_text.config(state=tk.DISABLED)

    def show_teachers_hours(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Часы преподавателей")
        dialog.geometry("850x600")
        dialog.transient(self.root)

        columns = ('Преподаватель', 'Всего часов', 'Кол-во занятий', 'Дни')
        tree = ttk.Treeview(dialog, columns=columns, show='headings', height=20)

        tree.heading('Преподаватель', text='Преподаватель')
        tree.heading('Всего часов', text='Всего часов')
        tree.heading('Кол-во занятий', text='Кол-во занятий')
        tree.heading('Дни', text='Дни')

        tree.column('Преподаватель', width=250)
        tree.column('Всего часов', width=120)
        tree.column('Кол-во занятий', width=120)
        tree.column('Дни', width=300)

        scrollbar = ttk.Scrollbar(dialog, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=10)

        teacher_stats = defaultdict(lambda: {'hours': 0, 'lessons': 0, 'days': set()})

        for group in self.schedule_data:
            for day in self.days_of_week:
                for lesson in self.schedule_data.get(group, {}).get(day, {}).values():
                    teachers = re.split(r'[/,]', lesson.get('преподаватель', ''))
                    hours = lesson.get('часы', 2)
                    for t in teachers:
                        t = t.strip()
                        if t:
                            teacher_stats[t]['hours'] += hours
                            teacher_stats[t]['lessons'] += 1
                            teacher_stats[t]['days'].add(day)

        total_hours = 0
        for teacher, stats in sorted(teacher_stats.items(), key=lambda x: x[1]['hours'], reverse=True):
            tree.insert('', tk.END, values=(
                teacher,
                stats['hours'],
                stats['lessons'],
                ', '.join(
                    sorted(stats['days'], key=lambda x: self.days_of_week.index(x) if x in self.days_of_week else 99))
            ))
            total_hours += stats['hours']

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(btn_frame, text="Экспорт в Excel",
                   command=lambda: self.export_teacher_stats_to_excel(teacher_stats)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Закрыть", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

        info_label = ttk.Label(btn_frame,
                               text=f"Всего преподавателей: {len(teacher_stats)}, Общее количество часов: {total_hours}",
                               font=('Arial', 10))
        info_label.pack(side=tk.LEFT, padx=20)

    def export_teacher_stats_to_excel(self, teacher_stats):
        file_path = filedialog.asksaveasfilename(
            title="Сохранить статистику преподавателей",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            rows = []
            for teacher, stats in sorted(teacher_stats.items(), key=lambda x: x[1]['hours'], reverse=True):
                rows.append({
                    'Преподаватель': teacher,
                    'Всего часов': stats['hours'],
                    'Количество занятий': stats['lessons'],
                    'Дни': ', '.join(sorted(stats['days'],
                                            key=lambda x: self.days_of_week.index(x) if x in self.days_of_week else 99))
                })

            if rows:
                df = pd.DataFrame(rows)
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Успех", f"Данные экспортированы в {file_path}")
            else:
                messagebox.showwarning("Внимание", "Нет данных для экспорта")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать данные: {str(e)}")

    def export_teachers_hours(self):
        teacher = self.teacher_var.get()
        if not teacher:
            messagebox.showwarning("Внимание", "Выберите преподавателя")
            return

        file_path = filedialog.asksaveasfilename(
            title=f"Сохранить расписание преподавателя {teacher}",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            rows = []
            total_hours = 0

            for group in self.schedule_data:
                for day in self.days_of_week:
                    for para, lesson in self.schedule_data.get(group, {}).get(day, {}).items():
                        teachers = re.split(r'[/,]', lesson.get('преподаватель', ''))
                        for t in teachers:
                            if t.strip() == teacher:
                                hours = lesson.get('часы', 2)
                                rows.append({
                                    'День': day,
                                    'Пара': para,
                                    'Группа': group,
                                    'Предмет': lesson.get('предмет', ''),
                                    'Кабинет': lesson.get('кабинет', ''),
                                    'Часы': hours
                                })
                                total_hours += hours

            if rows:
                day_order = {day: i for i, day in enumerate(self.days_of_week)}
                rows.sort(key=lambda x: (day_order.get(x['День'], 99), int(x['Пара']) if x['Пара'].isdigit() else 99))

                df = pd.DataFrame(rows)
                df.to_excel(file_path, index=False)

                book = load_workbook(file_path)
                ws = book.active
                last_row = ws.max_row + 2
                ws.cell(row=last_row, column=1, value="ИТОГО:")
                ws.cell(row=last_row, column=6, value=total_hours)
                ws.cell(row=last_row, column=6).font = Font(bold=True)
                book.save(file_path)

                messagebox.showinfo("Успех", f"Данные экспортированы в {file_path}")
            else:
                messagebox.showwarning("Внимание", f"Нет занятий у преподавателя {teacher}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать данные: {str(e)}")

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def delete_selected_lesson(self):
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.tree.item(item, 'values')
        para_num = values[0]

        if messagebox.askyesno("Подтверждение", f"Удалить пару {para_num}?"):
            if self.current_group in self.schedule_data:
                if self.current_day in self.schedule_data[self.current_group]:
                    if para_num in self.schedule_data[self.current_group][self.current_day]:
                        del self.schedule_data[self.current_group][self.current_day][para_num]
                        self.update_schedule_display()
                        self.update_editor_table()
                        self.save_to_excel()
                        self.update_teachers_list()

    def edit_selected_lesson(self):
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.tree.item(item, 'values')
        para_num = values[0]

        if self.current_group in self.schedule_data:
            if self.current_day in self.schedule_data[self.current_group]:
                if para_num in self.schedule_data[self.current_group][self.current_day]:
                    lesson = self.schedule_data[self.current_group][self.current_day][para_num]
                    self.edit_lesson_dialog(para_num, lesson)

    def edit_lesson_dialog(self, para_num, lesson_data):
        dialog = tk.Toplevel(self.root)
        dialog.title("Редактировать пару")
        dialog.geometry("550x400")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Номер пары:", font=('Arial', 10)).pack(pady=5)
        para_var = tk.StringVar(value=para_num)
        para_combo = ttk.Combobox(dialog, textvariable=para_var,
                                  values=['1', '2', '3', '4', '5', '6'], width=10)
        para_combo.pack(pady=5)

        ttk.Label(dialog, text="Предмет:", font=('Arial', 10)).pack(pady=5)
        subject_var = tk.StringVar(value=lesson_data.get('предмет', ''))
        subject_combo = ttk.Combobox(dialog, textvariable=subject_var,
                                     values=self.subjects_list, width=45)
        subject_combo.pack(pady=5)

        ttk.Label(dialog, text="Преподаватель:", font=('Arial', 10)).pack(pady=5)
        teacher_var = tk.StringVar(value=lesson_data.get('преподаватель', ''))
        teacher_combo = ttk.Combobox(dialog, textvariable=teacher_var,
                                     values=self.teachers_list, width=45)
        teacher_combo.pack(pady=5)

        ttk.Label(dialog, text="Кабинет:", font=('Arial', 10)).pack(pady=5)
        room_var = tk.StringVar(value=lesson_data.get('кабинет', ''))
        room_entry = ttk.Entry(dialog, textvariable=room_var, width=45)
        room_entry.pack(pady=5)

        ttk.Label(dialog, text="Часы (1 пара = 2 часа):", font=('Arial', 10)).pack(pady=5)
        hours_var = tk.StringVar(value=str(lesson_data.get('часы', 2)))
        hours_entry = ttk.Entry(dialog, textvariable=hours_var, width=10)
        hours_entry.pack(pady=5)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)

        def save_lesson():
            new_para = para_var.get()
            subject = subject_var.get().strip()
            teacher = teacher_var.get().strip()
            room = room_var.get().strip()
            hours = float(hours_var.get()) if hours_var.get() else 2

            if not new_para or not subject:
                messagebox.showwarning("Внимание", "Номер пары и предмет обязательны")
                return

            if self.current_group in self.schedule_data:
                if self.current_day in self.schedule_data[self.current_group]:
                    if para_num in self.schedule_data[self.current_group][self.current_day]:
                        del self.schedule_data[self.current_group][self.current_day][para_num]

            if self.current_group not in self.schedule_data:
                self.schedule_data[self.current_group] = {day: {} for day in self.days_of_week}
            if self.current_day not in self.schedule_data[self.current_group]:
                self.schedule_data[self.current_group][self.current_day] = {}

            self.schedule_data[self.current_group][self.current_day][new_para] = {
                'предмет': subject,
                'преподаватель': teacher,
                'кабинет': room,
                'часы': hours
            }

            dialog.destroy()
            self.update_schedule_display()
            self.update_editor_table()
            self.save_to_excel()
            self.update_teachers_list()

        ttk.Button(btn_frame, text="💾 Сохранить", command=save_lesson, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ Отмена", command=dialog.destroy, width=15).pack(side=tk.LEFT, padx=5)

    def on_group_change(self, event):
        self.current_group = self.group_var.get()
        self.update_schedule_display()
        if self.editor_group_var.get() != self.current_group:
            self.editor_group_var.set(self.current_group)
            self.update_editor_table()

    def on_day_change(self, event):
        self.current_day = self.day_var.get()
        self.update_schedule_display()

    def update_schedule_display(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not self.current_group or not self.current_day:
            return

        day_schedule = self.schedule_data.get(self.current_group, {}).get(self.current_day, {})

        for para_num in sorted(day_schedule.keys(), key=lambda x: int(x) if x.isdigit() else 0):
            lesson = day_schedule[para_num]
            self.tree.insert('', tk.END, values=(
                para_num,
                lesson.get('предмет', ''),
                lesson.get('преподаватель', ''),
                lesson.get('кабинет', ''),
                lesson.get('часы', 2)
            ))

    def add_lesson(self):
        if not self.current_group or not self.current_day:
            messagebox.showwarning("Внимание", "Выберите группу и день")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить пару")
        dialog.geometry("550x400")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Номер пары:", font=('Arial', 10)).pack(pady=5)
        para_var = tk.StringVar()
        para_combo = ttk.Combobox(dialog, textvariable=para_var,
                                  values=['1', '2', '3', '4', '5', '6'], width=10)
        para_combo.pack(pady=5)

        ttk.Label(dialog, text="Предмет:", font=('Arial', 10)).pack(pady=5)
        subject_var = tk.StringVar()
        subject_combo = ttk.Combobox(dialog, textvariable=subject_var,
                                     values=self.subjects_list, width=45)
        subject_combo.pack(pady=5)

        ttk.Label(dialog, text="Преподаватель:", font=('Arial', 10)).pack(pady=5)
        teacher_var = tk.StringVar()
        teacher_combo = ttk.Combobox(dialog, textvariable=teacher_var,
                                     values=self.teachers_list, width=45)
        teacher_combo.pack(pady=5)

        ttk.Label(dialog, text="Кабинет:", font=('Arial', 10)).pack(pady=5)
        room_var = tk.StringVar()
        room_entry = ttk.Entry(dialog, textvariable=room_var, width=45)
        room_entry.pack(pady=5)

        ttk.Label(dialog, text="Часы (1 пара = 2 часа):", font=('Arial', 10)).pack(pady=5)
        hours_var = tk.StringVar(value="2")
        hours_entry = ttk.Entry(dialog, textvariable=hours_var, width=10)
        hours_entry.pack(pady=5)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)

        def save_lesson():
            para = para_var.get()
            subject = subject_var.get().strip()
            teacher = teacher_var.get().strip()
            room = room_var.get().strip()
            hours = float(hours_var.get()) if hours_var.get() else 2

            if not para or not subject:
                messagebox.showwarning("Внимание", "Номер пары и предмет обязательны")
                return

            if self.current_group not in self.schedule_data:
                self.schedule_data[self.current_group] = {day: {} for day in self.days_of_week}
            if self.current_day not in self.schedule_data[self.current_group]:
                self.schedule_data[self.current_group][self.current_day] = {}

            self.schedule_data[self.current_group][self.current_day][para] = {
                'предмет': subject,
                'преподаватель': teacher,
                'кабинет': room,
                'часы': hours
            }

            dialog.destroy()
            self.update_schedule_display()
            self.update_editor_table()
            self.save_to_excel()
            self.update_teachers_list()

        ttk.Button(btn_frame, text="💾 Сохранить", command=save_lesson, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ Отмена", command=dialog.destroy, width=15).pack(side=tk.LEFT, padx=5)

    def import_from_excel(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл для импорта",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            temp_app = ScheduleApp.__new__(ScheduleApp)
            temp_app.schedule_data = {}
            temp_app.group_names = []
            temp_app.days_of_week = self.days_of_week

            success = temp_app.load_excel_data(file_path)

            if success:
                imported_count = 0
                new_groups = set()

                for group, days in temp_app.schedule_data.items():
                    if group not in self.schedule_data:
                        self.schedule_data[group] = {d: {} for d in self.days_of_week}
                        new_groups.add(group)

                    for day, lessons in days.items():
                        if day not in self.schedule_data[group]:
                            self.schedule_data[group][day] = {}
                        for para, lesson in lessons.items():
                            self.schedule_data[group][day][para] = lesson
                            imported_count += 1

                if new_groups:
                    self.group_names.extend(list(new_groups))
                    self.group_names = sorted(set(self.group_names))

                self.group_combo['values'] = self.group_names
                self.editor_group_combo['values'] = self.group_names
                if self.current_group in self.group_names:
                    self.group_var.set(self.current_group)
                    self.editor_group_var.set(self.current_group)
                elif self.group_names:
                    self.current_group = self.group_names[0]
                    self.group_var.set(self.current_group)
                    self.editor_group_var.set(self.current_group)

                self.save_to_excel()
                self.refresh_all()
                self.imported_files.append(file_path)

                messagebox.showinfo("Успех",
                                    f"Импортировано {imported_count} записей\n"
                                    f"Добавлено групп: {len(new_groups)}\n"
                                    f"Всего групп: {len(self.group_names)}")
            else:
                self.import_from_simple_excel(file_path)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать данные: {str(e)}")

    def import_from_simple_excel(self, file_path):
        try:
            df = pd.read_excel(file_path, header=0)

            required_cols = ['Группа', 'День', 'Пара', 'Предмет']
            missing_cols = [col for col in required_cols if col not in df.columns]

            if missing_cols:
                alt_cols = {
                    'Группа': ['Группа', 'группа', 'Group', 'group'],
                    'День': ['День', 'день', 'Day', 'day'],
                    'Пара': ['Пара', 'пара', 'Para', 'para'],
                    'Предмет': ['Предмет', 'предмет', 'Subject', 'subject']
                }

                for col in list(df.columns):
                    for key, alts in alt_cols.items():
                        if col in alts:
                            df.rename(columns={col: key}, inplace=True)

                missing_cols = [col for col in required_cols if col not in df.columns]

                if missing_cols:
                    messagebox.showerror("Ошибка",
                                         f"В файле не найдены необходимые колонки: {', '.join(missing_cols)}\n"
                                         f"Доступные колонки: {', '.join(df.columns)}")
                    return

            imported_count = 0
            new_groups = set()

            for _, row in df.iterrows():
                group = str(row['Группа']) if not pd.isna(row['Группа']) else ''
                day = str(row['День']) if not pd.isna(row['День']) else ''
                para = str(row['Пара']) if not pd.isna(row['Пара']) else ''
                subject = str(row['Предмет']) if not pd.isna(row['Предмет']) else ''
                teacher = str(row['Преподаватель']) if not pd.isna(row['Преподаватель']) else ''
                room = str(row['Кабинет']) if not pd.isna(row['Кабинет']) else ''
                hours = float(row['Часы']) if 'Часы' in df.columns and not pd.isna(row['Часы']) else 2

                if group and day and para and subject:
                    day_norm = self.normalize_day(day)
                    if not day_norm:
                        continue

                    if group not in self.schedule_data:
                        self.schedule_data[group] = {d: {} for d in self.days_of_week}
                        new_groups.add(group)

                    if day_norm not in self.schedule_data[group]:
                        self.schedule_data[group][day_norm] = {}

                    self.schedule_data[group][day_norm][para] = {
                        'предмет': subject,
                        'преподаватель': teacher,
                        'кабинет': room,
                        'часы': hours
                    }
                    imported_count += 1

            if new_groups:
                self.group_names.extend(list(new_groups))
                self.group_names = sorted(set(self.group_names))

            self.group_combo['values'] = self.group_names
            self.editor_group_combo['values'] = self.group_names
            if self.current_group in self.group_names:
                self.group_var.set(self.current_group)
                self.editor_group_var.set(self.current_group)
            elif self.group_names:
                self.current_group = self.group_names[0]
                self.group_var.set(self.current_group)
                self.editor_group_var.set(self.current_group)

            self.save_to_excel()
            self.refresh_all()
            self.imported_files.append(file_path)

            messagebox.showinfo("Успех",
                                f"Импортировано {imported_count} записей\n"
                                f"Добавлено групп: {len(new_groups)}\n"
                                f"Всего групп: {len(self.group_names)}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать данные: {str(e)}")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(
            title="Сохранить Excel файл",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if not file_path:
            return

        if self.save_to_excel(file_path):
            messagebox.showinfo("Успех", f"Данные успешно экспортированы в {file_path}")

    def refresh_data(self):
        file_path = 'общий массив.xlsx'
        if os.path.exists(file_path):
            if messagebox.askyesno("Подтверждение",
                                   "Обновить данные из исходного файла?\nТекущие изменения будут потеряны."):
                self.load_excel_data(file_path)
                self.save_to_excel()
                self.refresh_all()
                messagebox.showinfo("Успех", "Данные обновлены")
        elif os.path.exists('общий массив — копия.xlsx'):
            if messagebox.askyesno("Подтверждение", "Обновить данные из файла 'общий массив — копия.xlsx'?"):
                self.load_excel_data('общий массив — копия.xlsx')
                self.save_to_excel()
                self.refresh_all()
                messagebox.showinfo("Успех", "Данные обновлены")
        else:
            messagebox.showwarning("Внимание", "Исходные файлы не найдены")

    def save_current_schedule(self):
        if not self.current_group or not self.current_day:
            return

        date_key = datetime.now().strftime("%Y-%m-%d %H:%M")

        if date_key not in self.history_data:
            self.history_data[date_key] = {}

        if self.current_group not in self.history_data[date_key]:
            self.history_data[date_key][self.current_group] = {}

        self.history_data[date_key][self.current_group][self.current_day] = \
            self.schedule_data.get(self.current_group, {}).get(self.current_day, {})

        self.save_history()
        self.update_history_display()
        self.save_to_excel()
        messagebox.showinfo("Успех", "Расписание сохранено")

    def update_history_display(self):
        self.history_text.delete(1.0, tk.END)

        if not self.history_data:
            self.history_text.insert(tk.END, "История пуста\n")
            return

        for date, groups in sorted(self.history_data.items(), reverse=True)[:10]:
            self.history_text.insert(tk.END, f"\n📅 {date}\n", 'date')
            for group, days in groups.items():
                total_group_hours = 0
                total_lessons = 0
                self.history_text.insert(tk.END, f"  📚 {group}:\n", 'group')
                for day, lessons in days.items():
                    day_lessons = len(lessons)
                    day_hours = sum(l.get('часы', 2) for l in lessons.values())
                    total_lessons += day_lessons
                    total_group_hours += day_hours
                    self.history_text.insert(tk.END, f"    {day}: {day_lessons} пар, {day_hours} ч.\n", 'day')
                self.history_text.insert(tk.END, f"    Итого: {total_lessons} пар, {total_group_hours} ч.\n\n", 'total')

        self.history_text.tag_config('date', foreground='blue', font=('Arial', 10, 'bold'))
        self.history_text.tag_config('group', foreground='darkgreen', font=('Arial', 9, 'bold'))
        self.history_text.tag_config('day', foreground='purple')
        self.history_text.tag_config('total', foreground='red', font=('Arial', 9, 'bold'))

    def export_history(self):
        file_path = filedialog.asksaveasfilename(
            title="Сохранить историю",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            rows = []
            for date, groups in self.history_data.items():
                for group, days in groups.items():
                    for day, lessons in days.items():
                        for para, lesson in lessons.items():
                            rows.append({
                                'Дата сохранения': date,
                                'Группа': group,
                                'День': day,
                                'Пара': para,
                                'Предмет': lesson.get('предмет', ''),
                                'Преподаватель': lesson.get('преподаватель', ''),
                                'Кабинет': lesson.get('кабинет', ''),
                                'Часы': lesson.get('часы', 2)
                            })

            if rows:
                df = pd.DataFrame(rows)
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Успех", f"История экспортирована в {file_path}")
            else:
                messagebox.showwarning("Внимание", "История пуста")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать историю: {str(e)}")

    def show_hours_count(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Подсчет часов")
        dialog.geometry("750x550")
        dialog.transient(self.root)

        text = ScrolledText(dialog, width=80, height=25, font=('Arial', 10))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        text.insert(tk.END, "📊 ПОДСЧЕТ ЧАСОВ ПО ГРУППАМ И ДНЯМ\n", 'header')
        text.insert(tk.END, "=" * 60 + "\n\n", 'header')

        total_all_hours = 0
        total_all_lessons = 0

        for group in sorted(self.group_names):
            group_hours = 0
            group_lessons = 0
            text.insert(tk.END, f"📚 Группа {group}\n", 'group')
            text.insert(tk.END, "-" * 40 + "\n", 'group')

            for day in self.days_of_week:
                day_lessons = self.schedule_data.get(group, {}).get(day, {})
                day_count = len(day_lessons)
                day_hours = sum(l.get('часы', 2) for l in day_lessons.values())
                if day_count > 0:
                    text.insert(tk.END, f"  {day:12} | {day_count:2} пар | {day_hours:3} часов\n", 'day')
                    for para, lesson in sorted(day_lessons.items()):
                        text.insert(tk.END,
                                    f"    Пара {para}: {lesson.get('предмет', '')[:30]:30} "
                                    f"({lesson.get('часы', 2)} ч.)\n", 'lesson')
                group_hours += day_hours
                group_lessons += day_count

            text.insert(tk.END, f"\n  ИТОГО: {group_lessons} пар, {group_hours} часов\n", 'total')
            text.insert(tk.END, "=" * 40 + "\n\n", 'group')
            total_all_hours += group_hours
            total_all_lessons += group_lessons

        text.insert(tk.END, "=" * 60 + "\n", 'header')
        text.insert(tk.END, f"ОБЩИЕ ПОКАЗАТЕЛИ:\n", 'total_header')
        text.insert(tk.END, f"  Всего пар: {total_all_lessons}\n", 'total')
        text.insert(tk.END, f"  Всего часов: {total_all_hours}\n", 'total')
        text.insert(tk.END, f"  Количество групп: {len(self.group_names)}\n", 'total')
        if self.group_names:
            text.insert(tk.END, f"  Среднее часов на группу: {total_all_hours // len(self.group_names)}\n", 'total')

        text.tag_config('header', foreground='blue', font=('Arial', 14, 'bold'))
        text.tag_config('group', foreground='darkgreen', font=('Arial', 11, 'bold'))
        text.tag_config('day', foreground='black')
        text.tag_config('lesson', foreground='gray', font=('Arial', 9))
        text.tag_config('total', foreground='red', font=('Arial', 10, 'bold'))
        text.tag_config('total_header', foreground='red', font=('Arial', 12, 'bold'))

        text.config(state=tk.DISABLED)

        ttk.Button(dialog, text="Закрыть", command=dialog.destroy).pack(pady=10)

    def clear_history(self):
        if messagebox.askyesno("Подтверждение", "Очистить всю историю расписаний?"):
            self.history_data = {}
            self.save_history()
            self.update_history_display()
            messagebox.showinfo("Успех", "История очищена")


def main():
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()