import click
import os
from .logic import extract_vba, push_vba, start_watching # <-- NEW


@click.group()
def cli():
    """Ein CLI-Tool zum Bearbeiten von VBA-Code in modernen Editoren."""
    pass

@cli.command()
@click.argument('file', type=click.Path(exists=True, dir_okay=False))
@click.option('--output-dir', default='./vba_src', help='Verzeichnis für den extrahierten Code.')
def pull(file, output_dir):
    """Extrahiert VBA-Code aus einer Office-Datei."""
    click.echo(f"Extrahiere VBA-Code aus '{file}'...")
    try:
        extract_vba(file, output_dir)
        click.echo(f"✅ Extraktion erfolgreich. Code befindet sich in '{output_dir}'.")
    except Exception as e:
        click.echo(f"❌ Fehler bei der Extraktion: {e}")

@cli.command()
@click.argument('file', type=click.Path(exists=True, dir_okay=False))
@click.option('--source-dir', default='./vba_src', help='Verzeichnis mit dem zu importierenden Code.')
def push(file, source_dir):
    """Schreibt den lokalen Code zurück in die Office-Datei."""
    click.echo(f"Schreibe Code aus '{source_dir}' in '{file}'...")
    try:
        push_vba(source_dir, file)
        click.echo(f"✅ Push erfolgreich. Vergessen Sie nicht, die Backup-Datei '.bak' bei Bedarf zu löschen.")
    except Exception as e:
        click.echo(f"❌ Fehler beim Push: {e}")

@cli.command()
@click.argument('file', type=click.Path(exists=True, dir_okay=False))
@click.option('--source-dir', default='./vba_src', help='Directory to watch for changes.')
def watch(file, source_dir):
    """Monitors the source directory and auto-pushes on any change."""
    click.echo(f"🚀 Starting watcher for '{file}'. Press Ctrl+C to stop.")
    start_watching(source_dir, file)

if __name__ == '__main__':
    cli()