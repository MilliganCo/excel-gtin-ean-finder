#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


def analyze_file(path: str):
    # Читаем первый лист
    try:
        df = pd.read_excel(path)
    except Exception as e:
        raise RuntimeError(f"Не удалось прочитать Excel-файл: {e}")

    # Попробуем найти колонки по названию
    col_barcode = None
    col_name = None

    # Нормализуем имена столбцов для поиска
    cols_normalized = {str(c).strip().lower(): c for c in df.columns}

    # Для надёжности проверяем несколько вариантов написания
    for key in cols_normalized:
        if key in ["штрихкод", "штрих-код", "barcode"]:
            col_barcode = cols_normalized[key]
        if key in ["номенклатура", "товар", "наименование"]:
            col_name = cols_normalized[key]

    # Если не нашли, считаем, что первый столбец — Штрихкод, второй — Номенклатура
    if col_barcode is None or col_name is None:
        if df.shape[1] < 2:
            raise RuntimeError("В файле меньше двух столбцов, не из чего считать.")
        col_barcode = df.columns[0]
        col_name = df.columns[1]

    # Приводим колонки к понятным названиям (можно переименовать)
    df = df.rename(columns={
        col_barcode: "Штрихкод",
        col_name: "Номенклатура"
    })

    # Убедимся, что пустые строки не мешают
    df = df.dropna(subset=["Номенклатура", "Штрихкод"])

    # На всякий случай штрихкод как строка (чтобы не потерять ведущие нули)
    df["Штрихкод"] = df["Штрихкод"].astype(str).str.strip()
    df["Номенклатура"] = df["Номенклатура"].astype(str).str.strip()

    # Считаем количество разных штрихкодов на каждую номенклатуру
    counts = (
        df.groupby("Номенклатура")["Штрихкод"]
        .nunique()
        .rename("Количество_разных_штрихкодов")
    )

    df = df.join(counts, on="Номенклатура")

    # Статус: хорошо/плохо
    df["Статус"] = df["Количество_разных_штрихкодов"].apply(
        lambda x: "Только один штрихкод" if x == 1 else "ОК"
    )

    # Отдельная таблица с проблемными номенклатурами
    df_bad = (
        df[df["Количество_разных_штрихкодов"] == 1]
        .drop_duplicates(subset=["Номенклатура"])
        .loc[:, ["Номенклатура", "Количество_разных_штрихкодов"]]
        .sort_values("Номенклатура")
    )

    # Подготовка имени выходного файла
    folder = os.path.dirname(path)
    name = os.path.basename(path)
    root, ext = os.path.splitext(name)
    out_name = f"{root}_analyzed{ext}"
    out_path = os.path.join(folder, out_name)

    # Сохраняем в новый Excel
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Все данные с анализом", index=False)
        df_bad.to_excel(writer, sheet_name="Только 1 штрихкод", index=False)

    return out_path, len(df_bad)


def main():
    root = tk.Tk()
    root.withdraw()  # Не показывать пустое окно

    messagebox.showinfo(
        "Анализ штрихкодов",
        "Выберите Excel-файл (.xlsx), где первый столбец — Штрихкод, второй — Номенклатура."
    )

    file_path = filedialog.askopenfilename(
        title="Выберите Excel-файл",
        filetypes=[("Excel файлы", "*.xlsx;*.xls")]
    )

    if not file_path:
        # Пользователь отменил
        return

    try:
        out_path, bad_count = analyze_file(file_path)
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
        return

    messagebox.showinfo(
        "Готово",
        f"Анализ завершён.\n"
        f"Номенклатур только с одним штрихкодом: {bad_count}\n\n"
        f"Результат сохранён в:\n{out_path}"
    )


if __name__ == "__main__":
    # Для корректной работы Tkinter под Windows
    if getattr(sys, "frozen", False):
        # Если будете собирать в .exe через pyinstaller — иногда полезно
        os.chdir(sys._MEIPASS)
    main()
