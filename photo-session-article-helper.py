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
import threading
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import pasteboard
from exiftool import ExifToolHelper
from rich import print
from rich.prompt import Prompt
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


@dataclass
class Article:
    article_no: str
    article_desc: str
    collection: str
    color: str
    color_name: str
    article_categorie: str
    pos_front: str
    pos_back: str | None


class PhotoCreationHandler(FileSystemEventHandler):
    def __init__(self, target_file: str):
        self.target_file = target_file
        self.file_created = threading.Event()

    def on_created(self, event):
        if not event.is_directory and Path(event.src_path).name == self.target_file:
            self.file_created.set()


def read_excel_data(excel_file: Path):
    excel_data = {}
    found_header = False

    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active
    for (
        collection,
        article_no,
        article_desc,
        color,
        color_name,
        article_categorie,
        pos_front,
        pos_back,
    ) in sheet.rows:
        # ignore lines without data (header, heading, empty lines)
        if not found_header:
            if article_no.value is None:
                continue
            elif article_no.value == "ArtikelNr":
                found_header = True
                continue

        # read data
        excel_data[article_no.value] = Article(
            article_no.value,
            article_desc.value,
            collection.value,
            color.value,
            color_name.value,
            article_categorie.value,
            pos_front.value,
            pos_back.value,
        )

    print(f"- Anzahl an Artikeldaten im Excel: [dark_orange]{len(excel_data)}[/]")
    return excel_data


def read_article_data(excel_data) -> str | None:
    arcticle_no = Prompt.ask("[bold]ArtikelNr").strip()
    if not arcticle_no:
        return

    article = excel_data.get(arcticle_no)
    if article is None:
        print(f"[light_pink3]ArtikelNr '{arcticle_no}' nicht gefunden.")
        return

    return article


def read_position():
    while True:
        position = Prompt.ask("[bold]Position (v/h)").strip().lower()
        if not position:
            return

        if position not in ["v", "h"]:
            print("[light_pink3]Ungültige Position. Bitte 'v' oder 'h' eingeben.")
            continue

        return position


def set_clipboard_and_wait_for_photo(
    pb: pasteboard.Pasteboard, article: Article, position: str, watch_path: Path
):
    article_desc = article.article_desc.replace(".", "").replace(" ", "_")
    filename = f"{article.article_no}-{position}-{article.color}-{article_desc}.jpg"

    # Set up file system observer
    event_handler = PhotoCreationHandler(filename)
    observer = Observer()
    observer.schedule(event_handler, str(watch_path), recursive=False)
    observer.start()

    # Set clipboard content
    pb.set_contents(filename)
    print(
        f"[green]Filename [bold]'{pb.get_contents()}'[/] in die Zwischenablage kopiert."
    )

    # Wait for file creation or timeout
    try:
        result = event_handler.file_created.wait()
        if result:
            with ExifToolHelper() as et:
                et.set_tags(
                    watch_path / filename,
                    {
                        "IPTC:ObjectName": article.article_no,
                        "IPTC:Category": position,
                        "IPTC:Caption-Abstract": article.article_desc,
                        "IPTC:Headline": article.color,
                    },
                )
    finally:
        observer.stop()
        observer.join()


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

    print(f"- Suche Fotos in [dark_orange]{args.watch_path.absolute()}[/]")
    print(f"- Lese die Artikeldaten von [dark_orange]{args.excel_file.absolute()}[/]")
    return args


def main():
    args = parse_args()
    excel_data = read_excel_data(args.excel_file)
    pb = pasteboard.Pasteboard()

    try:
        while True:
            print("")
            print("=" * 80)
            article = read_article_data(excel_data)
            if article is None:
                break

            position = read_position()
            if position is None:
                break

            set_clipboard_and_wait_for_photo(pb, article, position, args.watch_path)

    except KeyboardInterrupt:
        ...


if __name__ == "__main__":
    print("\nStarte Foto Session")
    main()
    print("\nEnde der Foto Session")
