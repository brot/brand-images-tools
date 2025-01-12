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
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import pasteboard
from exiftool import ExifToolHelper
from rich.console import Console
from rich.live import Live
from rich.prompt import Prompt
from rich.table import Table
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

CONSOLE = Console()


@dataclass
class Article:
    article_no: str
    article_desc: str
    collection: str
    color: str
    color_name: str
    article_categorie: str
    position: str


class PhotoCreationHandler(FileSystemEventHandler):
    def __init__(self, target_file: str):
        self.target_file = target_file
        self.file_created = threading.Event()

    def on_created(self, event):
        if not event.is_directory and Path(event.src_path).name == self.target_file:
            self.file_created.set()


def read_excel_data(excel_file: Path):
    excel_data = defaultdict(list)
    found_header = False
    article_variations = 0

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
        positions = ["v", "h"] if pos_back.value == "x" else ["v"]
        for position in positions:
            article_variations += 1
            excel_data[article_no.value].append(
                Article(
                    article_no=article_no.value,
                    article_desc=article_desc.value,
                    article_categorie=article_categorie.value,
                    collection=collection.value,
                    color=color.value,
                    color_name=color_name.value,
                    position=position,
                )
            )

    CONSOLE.print(
        f"- Anzahl an Artikeldaten im Excel: [dark_orange]{len(excel_data)}[/]"
    )
    CONSOLE.print(
        f"- Anzahl an Artikelvariationen im Excel: [dark_orange]{article_variations}[/]"
    )
    return excel_data


def read_article_data(excel_data) -> list[Article] | None:
    arcticle_no = Prompt.ask("[bold]ArtikelNr").strip()
    if not arcticle_no:
        return

    article = excel_data.get(arcticle_no)
    if article is None:
        CONSOLE.print(f"[light_pink3]ArtikelNr '{arcticle_no}' nicht gefunden.")
        return

    return article


def generate_new_filename(article: Article, watch_path: Path) -> str:
    article_desc = article.article_desc.replace(".", "").replace(" ", "_")
    filename = (
        f"{article.article_no}-{article.position}-{article.color}-{article_desc}.jpg"
    )

    for i in itertools.count(1):
        if not (watch_path / filename).exists():
            break
        filename = f"{article.article_no}-{article.position}-{article.color}-{article_desc}-{i}.jpg"

    return filename


def set_clipboard_and_wait_for_photo(
    pb: pasteboard.Pasteboard, article: Article, watch_path: Path
):
    filename = generate_new_filename(article, watch_path)

    # Set clipboard content
    pb.set_contents(filename)
    CONSOLE.print(
        f"[green]Filename [bold]'{pb.get_contents()}'[/] in die Zwischenablage kopiert."
    )

    # Set up file system observer
    event_handler = PhotoCreationHandler(filename)
    observer = Observer()
    observer.schedule(event_handler, str(watch_path), recursive=False)
    observer.start()

    # Wait for file creation or timeout
    try:
        result = event_handler.file_created.wait()
        if result:
            with ExifToolHelper() as et:
                et.set_tags(
                    watch_path / filename,
                    {
                        "IPTC:ObjectName": article.article_no,
                        "IPTC:Category": article.position,
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

    CONSOLE.print(f"- Suche Fotos in [dark_orange]{args.watch_path.absolute()}[/]")
    CONSOLE.print(
        f"- Lese die Artikeldaten von [dark_orange]{args.excel_file.absolute()}[/]"
    )
    return args


def print_article_variations(articles: list[Article], selected_line: int = 0) -> Table:
    table = Table(
        title=f"Artikel {articles[0].article_no} hat {len(articles)} Variationen",
        row_styles=[
            "bold" if row_index == selected_line else ""
            for row_index in range(len(articles))
        ],
    )

    table.add_column("ArtikelNr")
    table.add_column("Artikelart")
    table.add_column("Artikelbezeichnung")
    table.add_column("Kollektion")
    table.add_column("Farbe")
    table.add_column("Position")
    table.add_column("Position")

    for article in articles:
        table.add_row(
            article.article_no,
            article.article_categorie,
            article.article_desc,
            article.collection,
            f"{article.color} - {article.color_name}",
            "vorne" if article.position == "v" else "hinten",
            article.position,
        )

    return table


def main():
    args = parse_args()
    excel_data = read_excel_data(args.excel_file)
    pb = pasteboard.Pasteboard()

    try:
        while True:
            CONSOLE.print("")
            CONSOLE.print("=" * 80)
            articles = read_article_data(excel_data)
            if articles is None:
                break

            with Live(
                print_article_variations(articles),
                auto_refresh=False,
                console=CONSOLE,
                transient=True,
            ) as live:
                for i, article in enumerate(articles):
                    live.update(print_article_variations(articles, i), refresh=True)
                    set_clipboard_and_wait_for_photo(pb, article, args.watch_path)
    except KeyboardInterrupt:
        ...


if __name__ == "__main__":
    CONSOLE.print("\nStarte Foto Session")
    main()
    CONSOLE.print("\nEnde der Foto Session")
