import os
from datetime import datetime
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Настройка страницы для удобного отображения на смартфоне
st.set_page_config(page_title="Журнал старосты", page_icon="🎓", layout="centered")

# Стили Excel
ЗЕЛЕНЫЙ = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
КРАСНЫЙ = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
ЖЕЛТЫЙ = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
ШАПКА = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
РАМКА = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def загрузить_книгу(группа):
    имя_файла = f"{группа}.xlsx"
    if os.path.exists(имя_файла):
        return openpyxl.load_workbook(имя_файла), имя_файла
    
    книга = openpyxl.Workbook()
    л1 = книга.active
    л1.title = "Посещаемость"
    л1.append(["№", "ФИО Студента"])
    
    л2 = книга.create_sheet(title="Успеваемость")
    л2.append(["№", "ФИО Студента"])
    
    for лист in [л1, л2]:
        лист.column_dimensions["A"].width = 5
        лист.column_dimensions["B"].width = 35
        for ячейка in лист:
            ячейка.font = Font(bold=True, size=11)
            ячейка.fill = ШАПКА
            ячейка.alignment = Alignment(horizontal="center")
            ячейка.border = РАМКА
    книга.save(имя_файла)
    return книга, имя_файла

st.title("🎓 Электронный журнал старосты")

# Выбор группы
группа = st.text_input("Введите название группы (например, ИС-21):", value="ИС-21").strip()

if группа:
    книга, имя_файла = загрузить_книгу(группа)
    лист_посещ = книга["Посещаемость"]
    лист_успев = книга["Успеваемость"]

    # Кнопка скачивания готового Excel-файла прямо в телефон
    with open(имя_файла, "rb") as file:
        st.download_button(
            label=f"📥 Скачать журнал {группа}.xlsx",
            data=file,
            file_name=имя_файла,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # Вкладки для мобильного экрана
    вкладка1, вкладка2, вкладка3, вкладка4 = st.tabs(["✅ Посещаемость", "⭐ Оценки", "📊 Сводка", "👥 Студенты"])

    # --- ВКЛАДКА 1: ПОСЕЩАЕМОСТЬ ---
    with вкладка1:
        if лист_посещ.max_row <= 1:
            st.warning("Сначала добавьте студентов во вкладке '👥 Студенты'.")
        else:
            дата = st.text_input("Дата занятия:", value=datetime.now().strftime("%d.%m"))
            st.write("Отметьте присутствующих:")
            
            отметки = {}
            for row in range(2, лист_посещ.max_row + 1):
                фио = лист_посещ.cell(row=row, column=2).value
                отметки[row] = st.radio(
                    f"{row - 1}. {фио}",
                    ["+", "н", "у"],
                    horizontal=True,
                    key=f"att_{row}"
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
                
                книга.save(имя_файла)
                st.success(f"Посещаемость за {дата} сохранена!")

    # --- ВКЛАДКА 2: ОЦЕНКИ ---
    with вкладка2:
        if лист_успев.max_row <= 1:
            st.warning("Сначала добавьте студентов во вкладке '👥 Студенты'.")
        else:
            предмет = st.text_input("Название предмета:", placeholder="Например: Информатика")
            студенты = [лист_успев.cell(row=r, column=2).value for r in range(2, лист_успев.max_row + 1)]
            выбранный_студент = st.selectbox("Выберите студента:", студенты)
            оценка = st.selectbox("Оценка:", ["5", "4", "3", "2", "зачет", "незачет"])
            
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
                    книга.save(имя_файла)
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
        новое_фио = st.text_input("ФИО нового студента:", placeholder="Иванов Иван")
        if st.button("Добавить в группу", use_container_width=True):
            if новое_фио.strip():
                for имя_л in ["Посещаемость", "Успеваемость"]:
                    л = книга[имя_л]
                    л.append([л.max_row, новое_фио.strip()])
                    for яч in л[л.max_row]:
                        яч.border = РАМКА
                книга.save(имя_файла)
                st.success(f"Студент {новое_фио} добавлен!")
                st.rerun()
        
        st.subheader("Текущий список:")
        for row in range(2, лист_посещ.max_row + 1):
            st.write(f"**{row - 1}.** {лист_посещ.cell(row=row, column=2).value}")