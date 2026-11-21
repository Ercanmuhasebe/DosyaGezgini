import json
import os
import pathlib
import sys


def configure(root):
    root = pathlib.Path(root).resolve()
    packages_dir = root / "packages"
    packages_dir.mkdir(parents=True, exist_ok=True)
    os.environ["ARGOS_PACKAGES_DIR"] = str(packages_dir)
    os.environ["ARGOS_DEVICE_TYPE"] = "cpu"
    os.environ["ARGOS_CHUNK_TYPE"] = "MINISBD"
    return root


def install_model(root, model_path):
    configure(root)
    import argostranslate.package

    argostranslate.package.install_from_path(pathlib.Path(model_path))


def translate(root, input_path, output_path):
    configure(root)
    import argostranslate.translate

    languages = argostranslate.translate.get_installed_languages()
    source = next((language for language in languages if language.code == "en"), None)
    target = next((language for language in languages if language.code == "tr"), None)
    if source is None or target is None:
        raise RuntimeError("English–Turkish Argos modeli kurulu değil.")
    translator = source.get_translation(target)
    if translator is None:
        raise RuntimeError("English–Turkish çeviri yolu bulunamadı.")

    with pathlib.Path(input_path).open("r", encoding="utf-8") as stream:
        rows = json.load(stream)

    translated = []
    for row in rows:
        title = row.get("title", "") or ""
        summary = row.get("summary", "") or ""
        translated.append({
            "title": translator.translate(title) if title.strip() else "",
            "summary": translator.translate(summary) if summary.strip() else "",
        })

    with pathlib.Path(output_path).open("w", encoding="utf-8") as stream:
        json.dump(translated, stream, ensure_ascii=False)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit("Kullanım: argos_translate.py install ROOT MODEL | translate ROOT INPUT OUTPUT")
    action = sys.argv[1]
    if action == "install" and len(sys.argv) == 4:
        install_model(sys.argv[2], sys.argv[3])
    elif action == "translate" and len(sys.argv) == 5:
        translate(sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        raise SystemExit("Geçersiz Argos komutu.")
