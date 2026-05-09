"""
Generate Unity .meta files for UPM/Git packages.

Unity Package Manager installs Git packages into an immutable cache; missing .meta
files cause "has no meta file, but it's in an immutable folder" and assets are ignored.

Run from repo root: python tools/generate_upm_metas.py
"""
from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def guid_for(rel_posix: str) -> str:
    return hashlib.md5(rel_posix.encode("utf-8")).hexdigest()


FOLDER_META = """fileFormatVersion: 2
guid: {guid}
folderAsset: yes
DefaultImporter:
  externalObjects: {{}}
  userData: 
  assetBundleName: 
  assetBundleVariant: 
"""

MONO_META = """fileFormatVersion: 2
guid: {guid}
MonoImporter:
  externalObjects: {{}}
  serializedVersion: 2
  defaultReferences: []
  executionOrder: 0
  icon: {{instanceID: 0}}
  userData: 
  assetBundleName: 
  assetBundleVariant: 
"""

ASMDEF_META = """fileFormatVersion: 2
guid: {guid}
AssemblyDefinitionImporter:
  externalObjects: {{}}
  userData: 
  assetBundleName: 
  assetBundleVariant: 
"""

PACKAGE_MANIFEST_META = """fileFormatVersion: 2
guid: {guid}
PackageManifestImporter:
  externalObjects: {{}}
  userData: 
  assetBundleName: 
  assetBundleVariant: 
"""

TEXT_META = """fileFormatVersion: 2
guid: {guid}
TextScriptImporter:
  externalObjects: {{}}
  userData: 
  assetBundleName: 
  assetBundleVariant: 
"""

DEFAULT_META = """fileFormatVersion: 2
guid: {guid}
DefaultImporter:
  externalObjects: {{}}
  userData: 
  assetBundleName: 
  assetBundleVariant: 
"""

PLUGIN_META = """fileFormatVersion: 2
guid: {guid}
PluginImporter:
  externalObjects: {{}}
  serializedVersion: 2
  iconMap: {{}}
  executionOrder: {{}}
  defineConstraints: []
  isPreloaded: 0
  isOverridable: 0
  isExplicitlyReferenced: 0
  validateReferences: 1
  platformData:
  - first:
      Any: 
    second:
      enabled: 0
      settings: {{}}
  - first:
      Editor: Editor
    second:
      enabled: 1
      settings:
        CPU: AnyCPU
        DefaultValueInitialized: true
        OS: AnyOS
  userData: 
  assetBundleName: 
  assetBundleVariant: 
"""


def git_ls_files() -> list[str]:
    r = subprocess.run(
        ["git", "ls-files"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=True,
    )
    return [line.strip() for line in r.stdout.splitlines() if line.strip()]


def meta_for_asset(rel_posix: str) -> str:
    ext = Path(rel_posix).suffix.lower()
    g = guid_for(rel_posix)
    if ext == ".cs":
        return MONO_META.format(guid=g)
    if ext == ".asmdef":
        return ASMDEF_META.format(guid=g)
    if rel_posix == "package.json":
        return PACKAGE_MANIFEST_META.format(guid=g)
    if ext == ".dll":
        return PLUGIN_META.format(guid=g)
    if ext in (".md", ".txt"):
        return TEXT_META.format(guid=g)
    return DEFAULT_META.format(guid=g)


def collect_dirs(files: list[str]) -> list[str]:
    found: set[str] = set()
    for f in files:
        parent = Path(f).parent
        while str(parent) != ".":
            found.add(str(parent).replace("\\", "/"))
            parent = parent.parent
    return sorted(found, key=lambda p: (p.count("/"), p.lower()))


def main() -> int:
    files = git_ls_files()
    written = 0

    for d in collect_dirs(files):
        meta_path = ROOT / f"{d}.meta"
        body = FOLDER_META.format(guid=guid_for(d))
        meta_path.write_text(body, encoding="utf-8", newline="\n")
        written += 1

    for rel in files:
        if rel.endswith(".meta"):
            continue
        meta_path = ROOT / f"{rel}.meta"
        body = meta_for_asset(rel)
        meta_path.write_text(body, encoding="utf-8", newline="\n")
        written += 1

    print(f"Wrote {written} .meta files under {ROOT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
