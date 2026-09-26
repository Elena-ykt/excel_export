import win32com.client
import os
import tkinter as tk
from tkinter import filedialog
from PIL import ImageGrab, Image

def export_excel_selection():
    # 1. Подключаемся к уже запущенному Excel
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        print("Ошибка: Excel не запущен. Откройте Excel и выделите нужный диапазон ячеек.")
        return

    wb = excel.ActiveWorkbook
    ws = excel.ActiveSheet
    sel = excel.Selection

    # 2. Проверяем, что выделен именно диапазон ячеек (а не график или фигура)
    try:
        _ = sel.Address
    except Exception:
        print("Ошибка: Выделите именно диапазон ячеек (таблицу), а не график или картинку.")
        return

    # 3. Инициализируем скрытое окно Tkinter для диалогов сохранения
    root = tk.Tk()
    root.withdraw() 

    # ================= ЭКСПОРТ В PDF =================
    pdf_path = filedialog.asksaveasfilename(
        title="Сохранить выделенный фрагмент как PDF",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")]
    )
    
    if not pdf_path:
        print("Экспорт в PDF отменен.")
        root.destroy()
        return

    # 0 = xlTypePDF. Экспортируем только выделенную область (sel)
    sel.ExportAsFixedFormat(0, pdf_path)
    print(f"✅ PDF успешно сохранен: {pdf_path}")

    # ================= ЭКСПОРТ В ИЗОБРАЖЕНИЕ =================
    img_path = filedialog.asksaveasfilename(
        title="Сохранить выделенный фрагмент как Изображение",
        defaultextension=".png",
        filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg")]
    )
    
    if not img_path:
        print("Экспорт в изображение отменен.")
        root.destroy()
        return

    # Копируем выделенную область как растровое изображение в буфер обмена
    # Appearance=1 (xlScreen - как на экране), Format=2 (xlBitmap - растровое изображение)
    # sel.CopyPicture(1, 2)
# все параметры (для справки):
# 1. Как на экране, растр (то, что было в скрипте)
# sel.CopyPicture(1, 2)
# → Чёткая картинка в экранном разрешении. Хорошо для PNG/JPG.

# 2. Как на экране, вектор
    # sel.CopyPicture(1, -4147)
# → Векторный метафайл. Идеально для вставки в Word/PowerPoint,
#   но Pillow может НЕ суметь его прочитать из буфера обмена.
# картинку не сохраняет - Ошибка: В буфере обмена нет изображения.
# Тип данных в буфере: <class 'NoneType'>

# 3. Как при печати, растр
# sel.CopyPicture(2, 2)
# → Растр, но с цветами/стилями, оптимизированными под печать.

# 4. Как при печати, вектор
    # sel.CopyPicture(2, -4147)
# → Вектор в «печатном» виде.
# картинку не сохраняет - Ошибка: В буфере обмена нет изображения.
# Тип данных в буфере: <class 'NoneType'>

    # Запоминаем текущий масштаб
    old_zoom = excel.ActiveWindow.Zoom

    # Увеличиваем масштаб (200% = в 2 раза больше пикселей)
    excel.ActiveWindow.Zoom = 200

    # Копируем
    sel.CopyPicture(1, 2)
    # Считываем изображение из буфера обмена
    img = ImageGrab.grabclipboard()

    # Возвращаем масштаб обратно
    excel.ActiveWindow.Zoom = old_zoom

    # ✅ Проверка типа: убеждаемся, что в буфере именно изображение
    if not isinstance(img, Image.Image):
        print("Ошибка: В буфере обмена нет изображения.")
        print(f"Тип данных в буфере: {type(img)}")
        root.destroy()
        return

    # Если сохраняем в JPG, нужно убрать альфа-канал (прозрачность)
    if img_path.lower().endswith(('.jpg', '.jpeg')) and img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')

    img.save(img_path)
    print(f"✅ Изображение успешно сохранено: {img_path}")

    root.destroy()

if __name__ == "__main__":
    print("Запустите скрипт, когда нужный фрагмент в Excel будет выделен.")
    export_excel_selection()