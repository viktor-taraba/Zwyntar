"""
Розбиває тексти пісень (.txt) на слова та зберігає результат в Excel.

Правила обробки:
  1. Весь текст приводиться до нижнього регістру.
  2. Фрагменти у квадратних дужках, напр. [Приспів], [Куплет 1], ігноруються.
  3. Назва пісні = назва файлу (без розширення .txt).
  4. Результат - .xlsx з двома колонками: "Назва пісні" та "Слово".
  5. Дублікати НЕ видаляються: слово, що повторюється 10 разів, дає 10 рядків.
  6. Розділові знаки (коми, крапки, ?, !, тире, лапки тощо) ігноруються.
     Дефіс усередині слова (по-різному, десь-там) зберігається, і слово
     лишається цілим. Дефіс/тире з пробілами навколо - це розділювач.

Використання:
    python lyrics_to_words.py                      # txt-файли з поточної папки
    python lyrics_to_words.py ./songs              # txt-файли з папки ./songs
    python lyrics_to_words.py ./songs -o words.xlsx
"""

import argparse
import re
import sys
import unicodedata
import pandas as pd
from pathlib import Path
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# [Приспів], [Куплет 1], [Інструментал] ... (може займати кілька рядків)
BRACKETS_RE = re.compile(r"\[[^\]]*\]", flags=re.DOTALL)

# Різні види апострофа приводимо до одного, щоб "кам'яна", "камʼяна", "кам’яна"
# вважалися одним і тим самим словом.
APOSTROPHES = "’ʼ‘`´′ˈ"
APOSTROPHE_TABLE = {ord(ch): "'" for ch in APOSTROPHES}

# Різні види дефіса (U+2010, U+2011) приводимо до звичайного "-".
# Тире (— та –) сюди не входять: вони завжди є розділювачами.
HYPHEN_TABLE = {ord("\u2010"): "-", ord("\u2011"): "-"}

# Слово = послідовність літер (будь-якої мови), яка може містити апостроф
# або дефіс ВСЕРЕДИНІ слова, між літерами (п'яти, по-різному, десь-там).
# Дефіс/тире, відокремлені пробілами (" - ", " — "), - це розділювач, а не слово.
WORD_RE = re.compile(r"[^\W\d_]+(?:['\-][^\W\d_]+)*")


def read_text(path: Path) -> str:
    """Читає файл, пробуючи кілька поширених кодувань."""
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    # Крайній випадок: читаємо, замінюючи невалідні байти
    return path.read_text(encoding="utf-8", errors="replace")


def extract_words(text: str) -> list[str]:
    """Повертає список слів тексту (з дублікатами) у нижньому регістрі."""
    text = unicodedata.normalize("NFC", text)
    text = BRACKETS_RE.sub(" ", text)          # прибираємо [Приспів] тощо
    text = text.translate(APOSTROPHE_TABLE)    # уніфікуємо апострофи
    text = text.translate(HYPHEN_TABLE)        # уніфікуємо дефіси
    text = text.lower()                        # нижній регістр
    return WORD_RE.findall(text)


def build_dataframe(folder: Path) -> pd.DataFrame:
    files = sorted(folder.glob("*.txt"))
    if not files:
        sys.exit(f"У папці '{folder}' не знайдено .txt файлів.")

    rows = []
    for file in files:
        song = file.stem  # назва пісні = назва файлу без розширення
        words = extract_words(read_text(file))
        rows.extend((song, w) for w in words)
        print(f"{song}: {len(words)} слів")

    return pd.DataFrame(rows, columns=["Назва пісні", "Слово"])


def save_excel(df: pd.DataFrame, output: Path) -> None:
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Слова")
        ws = writer.sheets["Слова"]

        # оформлення: шрифт, жирний заголовок, закріплений рядок, ширина колонок
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name="Arial", size=10, bold=(cell.row == 1))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for idx, col in enumerate(df.columns, start=1):
            max_len = max(df[col].astype(str).map(len).max(), len(col))
            ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 3, 60)


def main() -> None:
    parser = argparse.ArgumentParser(description="Розбиття текстів пісень на слова -> Excel")
    parser.add_argument("folder", nargs="?", default=".", help="папка з .txt файлами (за замовчуванням - поточна)")
    parser.add_argument("-o", "--output", default="songs_words.xlsx", help="ім'я вихідного .xlsx файлу")
    args = parser.parse_args()

    df = build_dataframe(Path(args.folder))
    save_excel(df, Path(args.output))
    print(f"\nГотово: {len(df)} рядків збережено у '{args.output}'")


if __name__ == "__main__":
    main()
