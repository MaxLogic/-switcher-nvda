from __future__ import annotations

import argparse
import re
from pathlib import Path
import zipfile


ROOT = Path(__file__).resolve().parent
ADDON_DIR = ROOT / "addon"
DIST_DIR = ROOT / "dist"
MANIFEST_PATH = ADDON_DIR / "manifest.ini"


def read_manifest_metadata():
	metadata = {}
	for line in MANIFEST_PATH.read_text(encoding="utf-8").splitlines():
		line = line.strip()
		if not line or line.startswith("#") or "=" not in line:
			continue
		key, value = (part.strip() for part in line.split("=", 1))
		metadata[key] = re.sub(r'^"|"$', "", value)
	return metadata


def iter_package_files():
	for path in ADDON_DIR.rglob("*"):
		if path.is_dir():
			continue
		if "__pycache__" in path.parts:
			continue
		if path.suffix in {".pyc", ".pyo"}:
			continue
		yield path


def build(output: Path | None = None) -> Path:
	DIST_DIR.mkdir(exist_ok=True)
	manifest = read_manifest_metadata()
	package_name = f'{manifest["name"]}-{manifest["version"]}.nvda-addon'
	package_path = output or (DIST_DIR / package_name)
	with zipfile.ZipFile(package_path, "w", zipfile.ZIP_DEFLATED) as bundle:
		for path in iter_package_files():
			bundle.write(path, path.relative_to(ADDON_DIR).as_posix())
	return package_path


def main() -> int:
	parser = argparse.ArgumentParser(description="Build the Voice Switcher NVDA add-on package.")
	parser.add_argument("-o", "--output", type=Path, help="Output .nvda-addon path.")
	args = parser.parse_args()
	package = build(args.output)
	print(package)
	return 0


if __name__ == "__main__":
	raise SystemExit(main())
