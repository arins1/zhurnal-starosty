import os
from datetime import datetime
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

st.set_page_config(page_title="Журнал колледжа", page_icon="🎓", layout="centered")

# Стили оформления Excel
ЗЕЛЕНЫЙ = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
КРАСНЫЙ = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
ЖЕЛТЫЙ = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
ШАПКА = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
РАМКА = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# Папка для хранения всех файлов групп
ПАПКА_ГРУПП = "groups_data"
os.makedirs(ПАПКА_ГРУПП, exist_ok=True)

def получить_список_групп():
    файлы = [f[:-5] for f in os.listdir(ПАПКА_ГРУПП) if f.endswith(".xlsx")]
    return sorted(файлы)

def создать_новый_журнал(название_группы):
    путь = os.path.join(ПАПКА_ГРУПП, f"{название_группы}.xlsx")
    if not os.path.exists(путь):
        книга = openpyxl.Workbook()
        л1 = книга.active
        л1.title = "Посещаемость"
        л1.append(["№", "ФИО Студента"])
        
        л2 = книга.create_sheet(title="Успеваемость")
        л2.append(["№", "ФИО Студента"])
        
        for лист in [л1, л2]:
            лист.column_dimensions["A"].width = 5
            лист.column_dimensions["B"].width = 35
            for ячейка in лист[1]:
                ячейка.font = Font(bold=True, size=11)
                ячейка.fill = ШАПКА
                ячейка.alignment = Alignment(horizontal="center")
                ячейка.border = РАМКА
        книга.save(путь)
    return путь

st.title("🎓 Электронный журнал колледжа")

# --- БОКОВАЯ ПАНЕЛЬ (УПРАВЛЕНИЕ ГРУППАМИ) ---
with st.sidebar:
    st.header("🏫 Управление группами")
    
    # Создание новой группы
    новая_группа = st.text_input("Создать новую группу:", placeholder="Например: ИС-21")
    if st.button("➕ Добавить группу", use_container_width=True):
        имя_чистое = новая_группа.strip().replace("/", "-").replace("\\", "-")
        if имя_чистое:
            создать_новый_журнал(имя_чистое)
            st.success(f"Группа {имя_чистое} создана!")
            st.rerun()
        else:
            st.warning("Введите название группы")

    st.divider()
    
    # Загрузка готового Excel-файла с телефона
    st.write("📤 **Загрузить готовый журнал (.xlsx):**")
    загруженный_файл = st.file_uploader("Выберите файл", type=["xlsx"], label_visibility="collapsed")
    if загруженный_файл is not None:
        путь_сохранения = os.path.join(ПАПКА_ГРУПП, загруженный_файл.name)
        with open(путь_сохранения, "wb") as f:
            f.write(загруженный_файл.getbuffer())
        st.success(f"Файл {загруженный_файл.name} загружен!")

# --- ОСНОВНОЙ ЭКРАН: ВЫБОР ГРУППЫ ---
список_групп = получить_список_групп()

if not список_групп:
    st.info("👋 Пока не создано ни одной группы. Создайте первую группу ниже или в меню слева (значок `>` в левом верхнем углу экрана).")
    старт_группа = st.text_input("Название вашей группы (например, 56-51):", placeholder="56-51")
    if st.button("Создать журнал группы", type="primary"):
        if старт_группа.strip():
            создать_новый_журнал(старт_группа.strip())
            st.rerun()
else:
    группа = st.selectbox("📌 Выберите учебную группу:", список_групп)
    имя_файла = f"{группа}.xlsx"
    путь_к_файлу = os.path.join(ПАПКА_ГРУПП, имя_файла)
    
    книга = openpyxl.load_workbook(путь_к_файлу)
    лист_посещ = книга["Посещаемость"]
    лист_успев = книга["Успеваемость"]

    # Кнопка скачивания Excel
    with open(путь_к_файлу, "rb") as file:
        st.download_button(
            label=f"📥 Скачать Excel-журнал ({имя_файла})",
            data=file,
            file_name=имя_файла,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    вкладка1, вкладка2, вкладка3, вкладка4 = st.tabs(["✅ Посещаемость", "⭐ Оценки", "📊 Сводка", "👥 Студенты"])

    # --- ВКЛАДКА 1: ПОСЕЩАЕМОСТЬ ---
    with вкладка1:
        if лист_посещ.max_row <= 1:
            st.warning("Сначала добавьте список студентов во вкладке '👥 Студенты'.")
        else:
            дата = st.text_input("Дата занятия:", value=datetime.now().strftime("%d.%m"))
            отметки = {}
            for row in range(2, лист_посещ.max_row + 1):
                фио = лист_посещ.cell(row=row, column=2).value
                отметки[row] = st.radio(
                    f"{row - 1}. {фио}",
                    ["+", "н", "у"],
                    horizontal=True,
                    key=f"att_{группа}_{row}"
                )
            
            if st.button("Сохранить посещаемость", type="primary", use_container_width=True):
                колонка = None
                for col in range(3, лист_посещ.max_column + 1):
                    if str(лист_посещ.cell(row=1, column=col).value) == дата:
                        колонка = col
                        break
                if колонка is None:
                    колонка = лист_посещ.max_column + 1
                    шапка = лист_посещ.cell(row=1, column=колонка, value=дата)
                    шапка.font = Font(bold=True)
                    шапка.fill = ШАПКА
                    шапка.alignment = Alignment(horizontal="center")
                    шапка.border = РАМКА
                
                for row, знач in отметки.items():
                    яч = лист_посещ.cell(row=row, column=колонка, value=знач)
                    яч.alignment = Alignment(horizontal="center")
                    яч.border = РАМКА
                    яч.fill = ЗЕЛЕНЫЙ if знач == "+" else (КРАСНЫЙ if знач == "н" else ЖЕЛТЫЙ)
                
                книга.save(путь_к_файлу)
                st.success(f"Посещаемость группы {группа} за {дата} сохранена!")

    # --- ВКЛАДКА 2: ОЦЕНКИ ---
    with вкладка2:
        if лист_успев.max_row <= 1:
            st.warning("Сначала добавьте студентов во вкладке '👥 Студенты'.")
        else:
            предмет = st.text_input("Название предмета:", placeholder="Например: Информатика", key=f"subj_{группа}")
            студенты = [лист_успев.cell(row=r, column=2).value for r in range(2, лист_успев.max_row + 1)]
            выбранный_студент = st.selectbox("Студент:", студенты, key=f"stud_{группа}")
            оценка = st.selectbox("Оценка:", ["5", "4", "3", "2", "зачет", "незачет"], key=f"grade_{группа}")
            
            if st.button("Поставить оценку", type="primary", use_container_width=True):
                if not предмет.strip():
                    st.error("Введите название предмета!")
                else:
                    колонка = None
                    for col in range(3, лист_успев.max_column + 1):
                        if str(лист_успев.cell(row=1, column=col).value).lower() == предмет.strip().lower():
                            колонка = col
                            break
                    if колонка is None:
                        колонка = лист_успев.max_column + 1
                        шапка = лист_успев.cell(row=1, column=колонка, value=предмет.strip())
                        шапка.font = Font(bold=True)
                        шапка.fill = ШАПКА
                        шапка.alignment = Alignment(horizontal="center")
                        шапка.border = РАМКА
                    
                    row_idx = студенты.index(выбранный_студент) + 2
                    яч = лист_успев.cell(row=row_idx, column=колонка)
                    яч.value = f"{яч.value}, {оценка}" if яч.value else оценка
                    яч.alignment = Alignment(horizontal="center")
                    яч.border = РАМКА
                    книга.save(путь_к_файлу)
                    st.success(f"Оценка '{оценка}' выставлена ({выбранный_студент})!")

    # --- ВКЛАДКА 3: СВОДКА ---
    with вкладка3:
        if лист_посещ.max_row <= 1:
            st.info("Нет данных для отображения.")
        else:
            данные_таблицы = []
            for row in range(2, лист_посещ.max_row + 1):
                фио = лист_посещ.cell(row=row, column=2).value
                всего, был, прогулы = 0, 0, 0
                for col in range(3, лист_посещ.max_column + 1):
                    з = лист_посещ.cell(row=row, column=col).value
                    if з:
                        всего += 1
                        if з == "+": был += 1
                        elif з == "н": прогулы += 1
                проц = f"{round((был / всего) * 100)}%" if всего > 0 else "0%"
                
                оценки = []
                for col in range(3, лист_успев.max_column + 1):
                    з = str(лист_успев.cell(row=row, column=col).value or "")
                    for ч in з.split(","):
                        if ч.strip().isdigit():
                            оценки.append(int(ч.strip()))
                ср = round(sum(оценки) / len(оценки), 2) if оценки else "-"
                
                данные_таблицы.append({
                    "ФИО": фио,
                    "Посещ.": проц,
                    "Прогулы": прогулы,
                    "Ср. балл": ср
                })
            st.dataframe(данные_таблицы, use_container_width=True)

    # --- ВКЛАДКА 4: СТУДЕНТЫ ---
    with вкладка4:
        st.write("Вставьте список студентов (каждый с новой строки):")
        список_текстом = st.text_area("ФИО студентов:", placeholder="Иванов Иван\nПетрова Анна\nСидоров Алексей", height=120, key=f"area_{группа}")
        if st.button("Добавить в группу", type="primary", use_container_width=True):
            строки = [s.strip() for s in список_текстом.split("\n") if s.strip()]
            if строки:
                for фио in строки:
                    for имя_л in ["Посещаемость", "Успеваемость"]:
                        л = книга[имя_л]
                        л.append([л.max_row, фио])
                        for яч in л[л.max_row]:
                            яч.border = РАМКА
                            книга.save(путь_к_файлу)
                            st.success(f"Добавлено студентов: {len(строки)}!")
                            st.rerun()
                            st.subheader(f"Состав группы {группа}:")
                            if лист_посещ.max_row <= 1:
                                st.write("Список пока пуст.")
                            else:
                                for row in range(2, лист_посещ.max_row + 1):
                                    st.write(f"{row - 1}. {лист_посещ.cell(row=row, column=2).value}")
