"""格式方案库：把"一整套设置"存成可命名、可勾选、可分享的单位。

一个方案 = 完整设置（`Settings.to_json()`）+ 名字 + 类型（期刊/学位）+ 来源。
内置两套（期刊论文默认 / 学位论文默认）只读，用户方案存成
`%APPDATA%\\论文格式助手\\schemes\\*.json`，一个方案一个文件。
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .ui.settings import Settings

APP_NAME = "论文格式助手"
KINDS = ("journal", "thesis")
KIND_LABEL = {"journal": "期刊论文", "thesis": "学位论文"}
BUILTIN_NAME = {"journal": "期刊论文默认", "thesis": "学位论文默认"}
SCHEME_VERSION = 1


def app_dir() -> Path:
    base = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA")
    return Path(base or Path.home()) / APP_NAME


def scheme_dir() -> Path:
    return app_dir() / "schemes"


def state_path() -> Path:
    return app_dir() / "settings.json"


@dataclass
class Scheme:
    name: str
    kind: str                      # journal | thesis
    settings: dict = field(default_factory=dict)
    source: str = ""               # 从哪个文件学来的
    builtin: bool = False
    updated: str = ""

    @property
    def label(self) -> str:
        return self.name + ("（内置）" if self.builtin else "")

    def to_json(self) -> dict:
        return {
            "version": SCHEME_VERSION,
            "name": self.name,
            "kind": self.kind,
            "source": self.source,
            "builtin": self.builtin,
            "updated": self.updated,
            "settings": self.settings,
        }

    @classmethod
    def from_json(cls, data: dict) -> "Scheme":
        return cls(
            name=str(data.get("name") or "未命名"),
            kind=str(data.get("kind") or "journal"),
            settings=dict(data.get("settings") or {}),
            source=str(data.get("source") or ""),
            builtin=bool(data.get("builtin")),
            updated=str(data.get("updated") or ""),
        )

    def settings_obj(self) -> Settings:
        s = Settings.from_json(self.settings)
        s.profile = self.kind
        return s


def _safe(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", name).strip("._") or "scheme"


def builtin(kind: str) -> Scheme:
    s = Settings()
    s._load_profile_preset(kind)
    s.profile = kind
    payload = s.scheme_payload()
    return Scheme(name=BUILTIN_NAME.get(kind, "内置方案"), kind=kind,
                  settings=payload, builtin=True)


def builtin_all() -> dict:
    return {k: builtin(k) for k in KINDS}


def load_user_schemes() -> list[Scheme]:
    """读用户方案。单个文件坏了就跳过，不让整个库打不开。"""
    out: list[Scheme] = []
    root = scheme_dir()
    if not root.is_dir():
        return out
    for path in sorted(root.glob("*.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            sc = Scheme.from_json(data)
            if sc.kind not in KINDS:
                continue
            sc.builtin = False
            out.append(sc)
        except Exception:      # noqa: BLE001
            continue
    return out


def find(name: str, kind: str | None = None) -> Scheme | None:
    """按名字找方案：内置优先，再找用户方案。"""
    for k in ((kind,) if kind else KINDS):
        b = builtin(k)
        if b.name == name:
            return b
    for sc in load_user_schemes():
        if sc.name == name and (kind is None or sc.kind == kind):
            return sc
    return None


def save(sc: Scheme) -> Path:
    """保存用户方案。内置方案不落盘（只读）。"""
    if sc.builtin:
        raise ValueError("内置方案不能保存")
    if not sc.name.strip():
        raise ValueError("方案名不能为空")
    root = scheme_dir()
    root.mkdir(parents=True, exist_ok=True)
    sc.updated = datetime.now().strftime("%Y-%m-%d %H:%M")
    path = root / f"{sc.kind}-{_safe(sc.name)}.json"
    path.write_text(json.dumps(sc.to_json(), ensure_ascii=False, indent=2),
                    encoding="utf-8")
    return path


def delete(sc: Scheme) -> bool:
    if sc.builtin:
        return False
    path = scheme_dir() / f"{sc.kind}-{_safe(sc.name)}.json"
    try:
        path.unlink()
        return True
    except OSError:
        return False


def rename(sc: Scheme, new_name: str) -> bool:
    if sc.builtin or not new_name.strip():
        return False
    old = scheme_dir() / f"{sc.kind}-{_safe(sc.name)}.json"
    new = scheme_dir() / f"{sc.kind}-{_safe(new_name)}.json"
    try:
        if new.exists() and new != old:
            return False
        sc.name = new_name
        sc.updated = datetime.now().strftime("%Y-%m-%d %H:%M")
        new.write_text(json.dumps(sc.to_json(), ensure_ascii=False, indent=2),
                       encoding="utf-8")
        if old.exists() and old != new:
            old.unlink()
        return True
    except OSError:
        return False


def move_kind(sc: Scheme, new_kind: str) -> bool:
    """把用户方案在「期刊论文 / 学位论文」之间移动。"""
    if sc.builtin or new_kind not in KINDS or new_kind == sc.kind:
        return False
    old = scheme_dir() / f"{sc.kind}-{_safe(sc.name)}.json"
    new = scheme_dir() / f"{new_kind}-{_safe(sc.name)}.json"
    try:
        if new.exists():
            return False
        sc.kind = new_kind
        sc.settings["profile"] = new_kind
        sc.updated = datetime.now().strftime("%Y-%m-%d %H:%M")
        new.write_text(json.dumps(sc.to_json(), ensure_ascii=False, indent=2),
                       encoding="utf-8")
        if old.exists() and old != new:
            old.unlink()
        return True
    except OSError:
        return False


def export(sc: Scheme, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(sc.to_json(), ensure_ascii=False, indent=2),
                    encoding="utf-8")
    return path


def import_file(path: Path) -> Scheme:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    sc = Scheme.from_json(data)
    if sc.kind not in KINDS:
        raise ValueError("方案类型不对")
    sc.builtin = False
    sc.name = sc.name.strip() or Path(path).stem
    return sc


def load_state() -> dict:
    try:
        return json.loads(state_path().read_text(encoding="utf-8"))
    except Exception:      # noqa: BLE001
        return {}


def save_state(last_name: str = "", last_kind: str = "") -> None:
    try:
        app_dir().mkdir(parents=True, exist_ok=True)
        state_path().write_text(
            json.dumps({"last_name": last_name, "last_kind": last_kind},
                       ensure_ascii=False, indent=2),
            encoding="utf-8")
    except Exception:      # noqa: BLE001
        pass
