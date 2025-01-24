# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "openpyxl",
#     "pasteboard",
#     "pyexiftool",
#     "rich",
#     "watchdog",
# ]
# ///

# How this script was initialized
#   uv init --script photo-session-article-helper.py --python 3.12
#   uv add --script photo-session-article-helper.py openpyxl pasteboard pyexiftool rich watchdog

import argparse
import itertools
import threading
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import pasteboard
from exiftool import ExifToolHelper
from exiftool.exceptions import ExifToolExecuteException
from rich.console import Console
from rich.prompt import Prompt
from rich.table import Table
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

CONSOLE = Console()
FILE_EXTENSION = ".NEF"


@dataclass
class Article:
    identity_no: str  # Identnummer
    sheet: str  # Tabellenblatt
    article_no: str  # ArtikelNr
    article_desc: str  # Artikelbezeichnung
    color_no: str  # Farbe


class PhotoCreationHandler(FileSystemEventHandler):
    def __init__(self, target_file: str):
        self.target_file = target_file
        self.file_created = threading.Event()

    def on_created(self, event):
        if not event.is_directory and Path(event.src_path).name == self.target_file:
            self.file_created.set()


def read_excel_data(excel_file: Path) -> dict[str, Article]:
    excel_data = {}
    found_header = False

    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active

    for sheet, identity_no, article_no, color_no, article_desc in sheet.rows:
        # ignore lines without data (header, heading, empty lines)
        if not found_header:
            if article_no.value is None:
                continue
            elif article_no.value == "ArtikelNr":
                found_header = True
                continue

        if identity_no.value in excel_data:
            raise ValueError(f"Identnummer '{article_no.value}' ist doppelt vorhanden!")

        # read data
        excel_data[identity_no.value] = Article(
            sheet=sheet.value,
            identity_no=identity_no.value,
            article_no=article_no.value,
            color_no=color_no.value,
            article_desc=article_desc.value,
        )

    CONSOLE.print(
        f"- Anzahl an Artikeldaten im Excel: [dark_orange]{len(excel_data)}[/]"
    )
    return excel_data


def ask_for_article_by_identity_no(excel_data) -> Article | None:
    while True:
        CONSOLE.print("")
        CONSOLE.print("=" * 70)

        identity_no = Prompt.ask("[bold]Identnummer").strip()
        if not identity_no:
            return

        article = excel_data.get(identity_no)
        if article is None:
            CONSOLE.print(f"[light_pink3]Identnummer '{identity_no}' nicht gefunden.")
            continue

        return article


def ask_for_next_action() -> str:
    while True:
        choice = Prompt.ask(
            "Drücke [bold]w[/]iederholen oder [bold]n[/]ächster Artikel",
            choices=["w", "n"],
        )
        return choice


def ask_for_side() -> str:
    return Prompt.ask("Vorder- oder Rückseite?", choices=["v", "r"])


def generate_new_filename(article: Article, side: str, watch_path: Path) -> Path:
    article_desc = article.article_desc.replace(".", "").replace(" ", "-")

    filename_template = "{article_no}_{color_no}_{article_desc}_{side}"
    filename_parts = {
        "article_no": article.article_no,
        "color_no": article.color_no,
        "article_desc": article_desc,
        "side": side,
    }
    filename = filename_template.format(**filename_parts)

    for i in itertools.count(1):
        filename = Path(filename).with_suffix(FILE_EXTENSION)
        if not (watch_path / filename).exists():
            break

        # if file already exists, add counter (starting with 1) ad the end of the filename
        filename = f"{filename_template}_{i}".format(**filename_parts)

    return filename


def set_clipboard_and_wait_for_photo(
    pb: pasteboard.Pasteboard, article: Article, watch_path: Path
):
    side = ask_for_side()
    filename = generate_new_filename(article, side, watch_path)

    # Set clipboard content
    pb.set_contents(filename.stem)
    CONSOLE.print(
        f"[green]Filename [bold]'{pb.get_contents()}'[/][/]"
    )

    # Set up file system observer
    event_handler = PhotoCreationHandler(str(filename))
    observer = Observer()
    observer.schedule(event_handler, str(watch_path), recursive=False)
    observer.start()

    # Wait for file creation or timeout
    try:
        result = event_handler.file_created.wait()
        if result:
            try:
                with ExifToolHelper() as et:
                    et.set_tags(
                        watch_path / filename,
                        {
                            "IPTC:ObjectName": article.article_no,
                            "IPTC:Category": side,
                            "IPTC:Caption-Abstract": article.article_desc,
                            "IPTC:Headline": article.color_no,
                        },
                    )
                CONSOLE.print(
                    f"[green]IPTC Daten von [bold]'{filename}'[/] erfolgreich aktualisiert[/]"
                )
            except ExifToolExecuteException as except_inst:
                CONSOLE.print(f"[light_pink3]{except_inst.stderr}[/]")
    finally:
        observer.stop()
        observer.join()


def print_article_info(article: Article) -> None:
    table = Table(
        "Identnummer",
        "ArtikelNr",
        "Farbe",
        "Artikelbezechnung",
    )

    table.add_row(
        article.identity_no,
        article.article_no,
        article.color_no,
        article.article_desc,
    )

    CONSOLE.print(table)


def process_article(article: Article, pb: pasteboard.Pasteboard, watch_path: Path):
    while True:
        set_clipboard_and_wait_for_photo(pb, article, watch_path)
        choice = ask_for_next_action()
        if choice == "n":
            break


def valid_path(path_str: str) -> Path:
    """Validate if path exists and return Path object."""
    path = Path(path_str)
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
    if not path.is_dir():
        raise argparse.ArgumentTypeError(
            f"Der Pfad '{path.absolute()}' ist kein Verzeichnis."
        )
    return path


def valid_file(path_str: str) -> Path:
    """Validate if file exists and return Path object."""
    filename = Path(path_str)
    if not filename.exists():
        raise argparse.ArgumentTypeError(
            f"Die Datei '{filename.absolute()}' existiert nicht."
        )
    if not filename.is_file():
        raise argparse.ArgumentTypeError(
            f"Die Datei '{filename.absolute()}' ist kein File."
        )
    return filename


def parse_args():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Verarbeitet Produktfotos mit Artikelnummern aus einer Excel-Datei"
    )
    parser.add_argument(
        "--excel",
        dest="excel_file",
        required=True,
        type=valid_file,
        help="Excel-Datei mit den Artikeldaten",
    )
    parser.add_argument(
        "--watch",
        dest="watch_path",
        type=valid_path,
        default=Path.home() / Path(".cache/brand-images/photos"),
        help="Verzeichnis, welches auf neue Fotos überwacht wird",
    )

    args = parser.parse_args()

    # Explicitly validate the watch_path whether it was provided or is using the default
    args.watch_path = valid_path(args.watch_path)

    CONSOLE.print(f"- Suche Fotos in [dark_orange]{args.watch_path.absolute()}[/]")
    CONSOLE.print(
        f"- Lese die Artikeldaten von [dark_orange]{args.excel_file.absolute()}[/]"
    )
    return args


def main():
    args = parse_args()
    excel_data = read_excel_data(args.excel_file)
    pb = pasteboard.Pasteboard()

    try:
        while True:
            article = ask_for_article_by_identity_no(excel_data)
            if article is None:
                break

            print_article_info(article)
            process_article(article, pb, args.watch_path)

    except KeyboardInterrupt:
        ...


if __name__ == "__main__":
    CONSOLE.print("\nStarte Foto Session")
    CONSOLE.print(
        f"- Unterstützes bzw. erwartetes Dateiformat: [dark_orange]{FILE_EXTENSION}[/]"
    )
    main()
    CONSOLE.print("\nEnde der Foto Session")
