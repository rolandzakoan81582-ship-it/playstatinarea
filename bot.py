import os
import asyncio
import shutil
import contextlib
from datetime import datetime
from dataclasses import dataclass
from io import BytesIO
import re
from typing import List, Dict, Optional, Tuple
from dotenv import load_dotenv

from aiogram import Bot, Dispatcher, F, Router
from aiogram.types import Message, CallbackQuery, FSInputFile, BufferedInputFile
from aiogram.exceptions import TelegramBadRequest
from aiogram.filters import CommandStart, Command, BaseFilter
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder

from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from aiohttp import ClientTimeout

# ---------- ENV ----------
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_IDS = {int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()}
EXCEL_PATH = os.getenv("EXCEL_PATH")
MASTER_SHEET = os.getenv("MASTER_SHEET", "Все игры")

AUTOSAVE_SECONDS = 1200  # 20 минут (секунды)

# ---------- Domain ----------
@dataclass
class AccountRow:
    row_idx: int
    login: str
    password: str

@dataclass
class GameSheet:
    name: str
    ws: Worksheet
    canonical_positions: Dict[str, List[int]]
    accounts: List[AccountRow]
    codes_col_idx: Optional[int] = None
    version_col_idx: Optional[int] = None
    note_col_idx: Optional[int] = None

class GameRepo:
    def __init__(self, path: str, master_sheet: str):
        self.path = path
        self.master_sheet = master_sheet
        self.wb = None
        self.games: Dict[str, GameSheet] = {}

    def load(self):
        if not os.path.exists(self.path):
            raise FileNotFoundError(f"Excel file not found: {self.path}")
        self.wb = load_workbook(self.path, data_only=False)
        self.games.clear()
        for sheet_name in self.wb.sheetnames:
            if sheet_name.strip() in (self.master_sheet.strip(), "История выдач"):
                continue
            ws = self.wb[sheet_name]
            headers = [cell.value.strip() if isinstance(cell.value, str) else cell.value for cell in ws[2]]
            canonical: Dict[str, List[int]] = {}
            codes_col_idx: Optional[int] = None
            version_col_idx: Optional[int] = None
            note_col_idx: Optional[int] = None
            for idx, h in enumerate(headers, start=1):
                if idx <= 2:
                    continue
                label = str(h).strip() if h is not None else ""
                if not label:
                    continue
                canon = self._canon_label(label)
                if canon:
                    canonical.setdefault(canon, []).append(idx)
                    continue
                llow = label.casefold()
                if "код" in llow:
                    codes_col_idx = idx
                if llow in ("версия", "издание", "edition"):
                    version_col_idx = idx
                if llow in ("пометка", "пометка менеджера", "заметка", "примечание", "note"):
                    note_col_idx = idx

            accounts: List[AccountRow] = []
            for r in range(3, ws.max_row + 1):
                login = ws.cell(r, 1).value
                password = ws.cell(r, 2).value
                if not login and not password:
                    continue
                accounts.append(AccountRow(r, str(login or "").strip(), str(password or "").strip()))

            gs = GameSheet(sheet_name, ws, canonical, accounts, codes_col_idx, version_col_idx, note_col_idx)
            self.games[sheet_name] = gs

            # Add dedicated metadata columns and migrate old annotations automatically.
            self.ensure_metadata_columns(sheet_name)
            for acc in accounts:
                current_version = ws.cell(acc.row_idx, gs.version_col_idx).value
                current_note = ws.cell(acc.row_idx, gs.note_col_idx).value
                inferred_version, inferred_note = self._parse_legacy_metadata(acc.password, sheet_name)
                login_version, login_note = self._parse_legacy_metadata(acc.login, sheet_name)
                # In a few legacy rows the manager annotation was appended to the email
                # instead of the password. Treat that as metadata too, without rewriting
                # the original login cell.
                if inferred_version == "Standard" and login_version != "Standard":
                    inferred_version = login_version
                combined_notes = " | ".join(x for x in (inferred_note, login_note) if x)
                if current_version is None or not str(current_version).strip() or str(current_version).strip().casefold() == "без версии":
                    ws.cell(acc.row_idx, gs.version_col_idx).value = inferred_version or "Standard"
                else:
                    ws.cell(acc.row_idx, gs.version_col_idx).value = self.canonicalize_version_name(sheet_name, str(current_version))
                if (current_note is None or not str(current_note).strip()) and combined_notes:
                    ws.cell(acc.row_idx, gs.note_col_idx).value = combined_notes[:500]

    @staticmethod
    def _canon_label(label: str) -> Optional[str]:
        s = label.replace("  ", " ").strip().upper()
        if "П2" in s and "PS4/PS5" in s:
            return "П2 PS4/PS5"
        if "П3" in s and "PS4" in s and "PS5" not in s:
            return "П3 PS4"
        if "П3" in s and "PS5" in s:
            return "П3 PS5"
        if s.startswith("П2") and "PS4" in s and "PS5" in s:
            return "П2 PS4/PS5"
        if s.startswith("П3") and "PS4" in s and "PS5" not in s:
            return "П3 PS4"
        if s.startswith("П3") and "PS5" in s:
            return "П3 PS5"
        return None

    def list_games(self) -> List[str]:
        return list(self.games.keys())

    def find_games_fuzzy(self, query: str, limit: int, score_cutoff: int) -> List[str]:
        """Case-insensitive fuzzy search."""
        choices = self.list_games()
        results = process.extract(
            query,
            choices,
            scorer=fuzz.WRatio,
            score_cutoff=score_cutoff,
            limit=limit,
            processor=lambda s: s.lower() if isinstance(s, str) else s
        )
        return [name for name, score, _ in results]

    @staticmethod
    def _normalize_edition_token(token: str) -> str:
        """Normalize common Latin/Cyrillic typos and abbreviations in edition labels."""
        raw = re.sub(r"[^a-zа-яё0-9]", "", str(token or "").casefold())
        aliases = {
            "Ultimate": ["ultimate", "ult", "ulti", "ultim", "ultimat", "ultimtae", "ultiamte", "ultiamte", "ultmate", "уьлтим", "уьлтимейт", "уьлт", "уьтим", "ультим", "ультимейт", "ульт", "ултимейт", "ултим", "ультимет"],
            "Deluxe": ["deluxe", "delux", "delixe", "delixe", "delixe", "delixe", "делюкс", "делюкс", "делюкс"],
            "Gold": ["gold", "голд", "гоld"],
            "Premium": ["premium", "prem", "премиум", "прем"],
            "Standard": ["standard", "standart", "standerd", "standartd", "стандарт", "стандард", "станд"],
            "Extra": ["extra", "экстра", "екстра", "экст", "екстр"],
            "Essential": ["essential", "essentia", "essent", "есентиал", "ессентиал", "эссеншл"],
            "Legendary": ["legendary", "легендари", "легендарное"],
            "Complete": ["complete", "полное", "полный", "полноеиздание", "polnoe", "polnoeizdanie"],
            "Definitive": ["definitive", "дефинитив"],
            "Collection": ["collection", "коллекшн", "коллекция"],
            "Anniversary": ["anniversary", "юбилейное"],
            "Champion": ["champion", "чемпион"],
            "Royal": ["royal", "роял"],
            "Devout": ["devout"],
            "Bundle": ["bundle", "бандл"],
            "GOTY": ["goty"],
            "Enhanced": ["enhanced", "enchanced", "улучшенное"],
            "Director's Cut": ["directorscut", "directorcut", "режиссерскаяверсия"],
            "American Wilds": ["americanwilds"],
        }
        # Exact aliases first; fuzzy matching is restricted to sufficiently long tokens.
        for canonical, variants in aliases.items():
            if raw in variants:
                return canonical
        if len(raw) >= 4:
            from rapidfuzz import fuzz
            best_name, best_score = "", 0
            for canonical, variants in aliases.items():
                for alias in variants:
                    if len(alias) >= 4:
                        score = fuzz.ratio(raw, alias)
                        if score > best_score:
                            best_name, best_score = canonical, score
            if best_score >= 82:
                return best_name
        return ""

    @classmethod
    def _parse_legacy_metadata(cls, password: str, game: str = "") -> Tuple[str, str]:
        """Return (edition, manager note) from a legacy password suffix.

        The original password cell is intentionally never rewritten: migration is
        additive and reversible. Any non-edition suffix is copied to the note field.
        """
        value = str(password or "").strip()
        if not value:
            return "Standard", ""
        # Passwords in this workbook are conventionally the first whitespace-delimited
        # token; trailing text, parentheses or separators are legacy annotations.
        match = re.match(r"^(\S+)(.*)$", value)
        if not match:
            return "Standard", ""
        tail = (match.group(2) or "").strip()
        if not tail:
            return "Standard", ""
        # Some subscription rows store several credential fields separated by pipes.
        # These are not edition labels or manager notes; preserve them verbatim.
        if tail.count("|") >= 2:
            return "Standard", ""
        # Remove surrounding punctuation while retaining readable words.
        cleaned = tail.strip(" |—-:;()[]{}")
        tokens = re.findall(r"[A-Za-zА-Яа-яЁё0-9]+", cleaned)
        edition = ""
        matched_tokens = set()
        # Multiword edition labels first.
        low = cleaned.casefold()
        multi = [
            (r"digital\s+deluxe", "Digital Deluxe"),
            (r"legendary\s+edition", "Legendary"),
            (r"anniversary\s+edition", "Anniversary"),
            (r"complete\s+edition", "Complete"),
            (r"definitive\s+edition", "Definitive"),
            (r"american\s+wilds(?:\s+edition)?", "American Wilds"),
            (r"director['’]?s\s+cut", "Director's Cut"),
        ]
        for pattern, label in multi:
            found = re.search(pattern, low)
            # "не American Wilds" is a warning that the account is NOT that edition.
            if found and label == "American Wilds" and re.search(r"не\s+$", low[:found.start()]):
                continue
            if found:
                edition = label
                matched_tokens.update(re.findall(r"[a-zа-яё]+", found.group(0)))
                break
        if not edition:
            for token in tokens:
                normalized = cls._normalize_edition_token(token)
                # Extra and Essential are PlayStation Plus subscription tiers; in ordinary
                # game sheets they are more often notes about a subscription, not game editions.
                game_low = str(game or "").casefold()
                if normalized in ("Extra", "Essential") and not ("подписк" in game_low or "ps plus" in game_low):
                    normalized = ""
                if normalized:
                    edition = normalized
                    matched_tokens.add(token.casefold())
                    break
        # PS platform, DLC, EA PLAY and operational remarks are notes, not editions.
        note_parts = []
        if edition:
            # Remove only the edition marker; preserve any accompanying manager text.
            note_text = cleaned
            for token in sorted(matched_tokens, key=len, reverse=True):
                note_text = re.sub(r"(?<!\w)" + re.escape(token) + r"(?!\w)", " ", note_text, flags=re.IGNORECASE)
            note_text = re.sub(r"\b(?:edition|издание|версия)\b", " ", note_text, flags=re.IGNORECASE)
            note_text = re.sub(r"[()|—,:;\-]+", " ", note_text)
            note_text = re.sub(r"\s+", " ", note_text).strip()
            if note_text:
                note_parts.append(note_text)
        else:
            note_parts.append(cleaned)
        note = " | ".join(x for x in note_parts if x)[:500]
        return edition or "Standard", note

    @staticmethod
    def _infer_legacy_version(password: str) -> str:
        return GameRepo._parse_legacy_metadata(password)[0]

    def ensure_metadata_columns(self, game: str) -> Tuple[int, int]:
        gs = self.games[game]
        if gs.version_col_idx is None:
            col = gs.ws.max_column + 1
            gs.ws.cell(2, col).value = "Версия"
            gs.version_col_idx = col
        if gs.note_col_idx is None:
            col = gs.ws.max_column + 1
            gs.ws.cell(2, col).value = "Пометка менеджера"
            gs.note_col_idx = col
        return gs.version_col_idx, gs.note_col_idx

    def get_account_version(self, game: str, row_idx: int) -> str:
        gs = self.games[game]
        if gs.version_col_idx:
            value = gs.ws.cell(row_idx, gs.version_col_idx).value
            if value is not None and str(value).strip():
                return str(value).strip()
        return "Standard"

    def canonicalize_version_name(self, game: str, value: str) -> str:
        raw = str(value or "").strip()
        if not raw:
            return "Standard"
        if re.fullmatch(r"(?:standard|standart|standerd|стандарт|стандард|станд)(?:\s+edition)?", raw, flags=re.IGNORECASE):
            return "Standard"
        parsed, _ = self._parse_legacy_metadata("x " + raw, game)
        # Use parsed edition only when it is a recognized alias; otherwise retain the
        # manager-defined custom name (for example, a special regional edition).
        if parsed != "Standard" or self._normalize_edition_token(raw):
            return parsed
        return raw

    def get_account_password(self, game: str, row_idx: int) -> str:
        raw = str(self.games[game].ws.cell(row_idx, 2).value or "").strip()
        match = re.match(r"^(\S+)(.*)$", raw)
        if not match or not (match.group(2) or "").strip():
            return raw
        tail = (match.group(2) or "").strip()
        # One legacy cell contains several pipe-separated credential/code fields.
        # Keep such a value intact rather than treating it as an annotation.
        if tail.count("|") >= 2:
            return raw
        return match.group(1)

    def get_versions(self, game: str) -> List[str]:
        gs = self.games[game]
        versions = []
        seen = set()
        for acc in gs.accounts:
            value = self.get_account_version(game, acc.row_idx) or "Standard"
            key = value.casefold()
            if key not in seen:
                seen.add(key)
                versions.append(value)
        # Standard should always be available for new stock, even if no Standard row exists yet.
        versions = [v for v in versions if v.casefold() != "standard"]
        versions.insert(0, "Standard")
        return versions

    def ensure_version_column(self, game: str) -> int:
        return self.ensure_metadata_columns(game)[0]

    def set_account_version(self, game: str, row_idx: int, version: str):
        col, _ = self.ensure_metadata_columns(game)
        self.games[game].ws.cell(row_idx, col).value = self.canonicalize_version_name(game, version or "Standard")

    def get_account_note(self, game: str, row_idx: int) -> str:
        gs = self.games[game]
        if not gs.note_col_idx:
            return ""
        value = gs.ws.cell(row_idx, gs.note_col_idx).value
        return str(value).strip() if value is not None else ""

    def set_account_note(self, game: str, row_idx: int, note: str):
        _, col = self.ensure_metadata_columns(game)
        self.games[game].ws.cell(row_idx, col).value = (note or "").strip() or None

    def iter_free_accounts(self, game: str, canon_position: str, version: Optional[str] = None):
        gs = self.games[game]
        cols = gs.canonical_positions.get(canon_position, [])
        for acc in gs.accounts:
            if version is not None:
                actual = self.get_account_version(game, acc.row_idx) or "Standard"
                if actual.casefold() != version.casefold():
                    continue
            chosen_col = None
            for c in cols:
                val = gs.ws.cell(acc.row_idx, c).value
                if isinstance(val, str) and val.strip() == "+":
                    chosen_col = c
                    break
            if chosen_col:
                yield acc, chosen_col

    def remove_plus(self, game: str, row_idx: int, col_idx: int):
        gs = self.games[game]
        cell = gs.ws.cell(row_idx, col_idx)
        if isinstance(cell.value, str) and cell.value.strip() == "+":
            cell.value = ""  # НЕ сохраняем сразу

    # ---- Search helpers ----
    @staticmethod
    def normalize_login_field(value: str) -> str:
        """Return the email part when old rows have a manager note appended to login."""
        raw = str(value or "").strip()
        if not raw:
            return ""
        tokens = raw.split()
        candidate = tokens[0].strip("`<>()[]{}.,;:")
        # Repair the rare legacy typo where a space was inserted after @.
        if candidate.endswith("@") and len(tokens) > 1:
            candidate += tokens[1].strip("`<>()[]{}.,;:")
        return candidate.casefold()

    def get_account_login(self, game: str, row_idx: int) -> str:
        raw = str(self.games[game].ws.cell(row_idx, 1).value or "").strip()
        if not raw:
            return ""
        tokens = raw.split()
        candidate = tokens[0].strip("`<>()[]{}.,;:")
        if candidate.endswith("@") and len(tokens) > 1:
            candidate += tokens[1].strip("`<>()[]{}.,;:")
        return candidate

    def find_account_by_login(self, login_query: str) -> Optional[Tuple[str, AccountRow]]:
        q = self.normalize_login_field(login_query)
        for game, gs in self.games.items():
            for acc in gs.accounts:
                if self.normalize_login_field(acc.login) == q:
                    return game, acc
        return None

    def _get_cols(self, game: str, canon: str) -> List[int]:
        return self.games[game].canonical_positions.get(canon, [])

    def cell_has_plus(self, game: str, row_idx: int, col: int) -> bool:
        gs = self.games[game]
        val = gs.ws.cell(row_idx, col).value
        return isinstance(val, str) and val.strip() == "+"

    def set_cell_plus(self, game: str, row_idx: int, col: int, present: bool) -> bool:
        gs = self.games[game]
        cell = gs.ws.cell(row_idx, col)
        old = isinstance(cell.value, str) and cell.value.strip() == "+"
        if present and not old:
            cell.value = "+"
            return True
        if (not present) and old:
            cell.value = ""
            return True
        return False

    def set_login(self, game: str, row_idx: int, new_login: str):
        gs = self.games[game]
        gs.ws.cell(row_idx, 1).value = new_login

    def set_password(self, game: str, row_idx: int, new_password: str):
        gs = self.games[game]
        gs.ws.cell(row_idx, 2).value = new_password

    def ensure_codes_column(self, game: str) -> int:
        gs = self.games[game]
        if gs.codes_col_idx is not None:
            return gs.codes_col_idx
        ws = gs.ws
        col = ws.max_column + 1
        ws.cell(2, col).value = "Коды"
        gs.codes_col_idx = col
        return col

    def get_codes(self, game: str, row_idx: int) -> List[str]:
        gs = self.games[game]
        c = gs.codes_col_idx
        if c is None:
            return []
        val = gs.ws.cell(row_idx, c).value
        if val is None or (isinstance(val, str) and not val.strip()):
            return []
        s = str(val).replace("\r\n", "\n").split("\n")
        return [x.strip() for x in s if x.strip()]

    def set_codes(self, game: str, row_idx: int, codes: List[str]) -> None:
        self.ensure_codes_column(game)
        gs = self.games[game]
        c = gs.codes_col_idx
        if not codes:
            gs.ws.cell(row_idx, c).value = None
        else:
            gs.ws.cell(row_idx, c).value = "\n".join(codes)

    def pop_first_code(self, game: str, row_idx: int) -> Optional[str]:
        codes = self.get_codes(game, row_idx)
        if not codes:
            return None
        self.set_codes(game, row_idx, codes[1:])
        return codes[0]

    def add_account_next_row(self, game: str, login: str, password: str, codes: Optional[List[str]] = None, version: str = "") -> int:
        gs = self.games[game]
        target_row = None
        for r in range(3, gs.ws.max_row + 2):
            lval = gs.ws.cell(r, 1).value
            pval = gs.ws.cell(r, 2).value
            if (lval is None or str(lval).strip() == "") and (pval is None or str(pval).strip() == ""):
                target_row = r
                break
        if target_row is None:
            target_row = gs.ws.max_row + 1

        gs.ws.cell(target_row, 1).value = login
        gs.ws.cell(target_row, 2).value = password
        for cols in gs.canonical_positions.values():
            for c in cols:
                gs.ws.cell(target_row, c).value = "+"

        gs.accounts.append(AccountRow(target_row, login, password))
        # Store even the explicit "Standard" choice so a password suffix
        # cannot accidentally override the edition selected in the bot.
        self.set_account_version(game, target_row, version or "Standard")
        if codes:
            self.set_codes(game, target_row, [c.strip() for c in codes if c and str(c).strip()])
        return target_row

    def find_accounts_by_login(self, login_query: str) -> List[Tuple[str, AccountRow]]:
        q = self.normalize_login_field(login_query)
        if not q:
            return []
        found = []
        for game, gs in self.games.items():
            for acc in gs.accounts:
                if self.normalize_login_field(acc.login) == q:
                    found.append((game, acc))
        return found

    def append_issue_history(self, game: str, row_idx: int, version: str, position: str, login: str, issued_by: int):
        title = "История выдач"
        if title not in self.wb.sheetnames:
            ws = self.wb.create_sheet(title)
            ws.append(["Дата и время", "Игра", "Версия", "Логин", "Позиция", "Выдал (Telegram ID)", "Пометка менеджера"])
        ws = self.wb[title]
        if ws.cell(1, 7).value is None:
            ws.cell(1, 7).value = "Пометка менеджера"
        ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), game, version or "Standard", login, position, issued_by, self.get_account_note(game, row_idx)])

    def delete_account_row(self, game: str, row_idx: int) -> bool:
        """Удалить строку аккаунта с листа и обновить список в памяти."""
        gs = self.games[game]
        if not any(acc.row_idx == row_idx for acc in gs.accounts):
            return False
        gs.ws.delete_rows(row_idx)
        accounts: List[AccountRow] = []
        for r in range(3, gs.ws.max_row + 1):
            login = gs.ws.cell(r, 1).value
            password = gs.ws.cell(r, 2).value
            if not login and not password:
                continue
            accounts.append(AccountRow(r, str(login or "").strip(), str(password or "").strip()))
        gs.accounts = accounts
        return True

    def save(self) -> bool:
        """Сохранить книгу на диск."""
        if not self.wb:
            return False
        try:
            self.wb.save(self.path)
            return True
        except Exception:
            return False

# ---------- FSM ----------
class IssueStates(StatesGroup):
    ChoosingGame = State()
    ChoosingVersion = State()
    ChoosingPosition = State()
    Offering = State()
    StealDecision = State()
    StealOffering = State()
    StealWrapConfirm = State()
    StealAwaitJump = State()
    IssuedCode = State()

class SearchStates(StatesGroup):
    Entry = State()
    WaitingLogin = State()
    AwaitDeleteLogin = State()
    Showing = State()
    EditMenu = State()
    AwaitLogin = State()
    AwaitPassword = State()
    AwaitCodes = State()
    AwaitNote = State()

class ExcelStates(StatesGroup):
    WaitingUpload = State()

class AddAccStates(StatesGroup):
    ChoosingGame = State()
    ChoosingEdition = State()
    WaitingNewEdition = State()
    WaitingBulkCreds = State()
    AskRemove = State()

class AuditStates(StatesGroup):
    WaitingLogins = State()

# ---------- UI ----------
def main_menu_kb():
    kb = ReplyKeyboardBuilder()
    kb.button(text="🎮 Выдать игру")
    kb.button(text="📤 Загрузить Excel")
    kb.button(text="🔍 Поиск аккаунта")
    kb.button(text="➕ Добавить аккаунт")
    kb.button(text="🧾 Сверить аккаунты")
    kb.button(text="💾 Сохранить изменения")
    kb.adjust(2, 2, 2)
    return kb.as_markup(resize_keyboard=True)

def back_inline_kb(extra: List[Tuple[str, str]] = None):
    kb = InlineKeyboardBuilder()
    if extra:
        for text, data in extra:
            kb.button(text=text, callback_data=data)
    kb.button(text="⬅️ Назад", callback_data="back")
    return kb.as_markup()

def positions_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="П2 PS4/PS5", callback_data="pos:P2")
    kb.button(text="П3 PS4", callback_data="pos:P3PS4")
    kb.button(text="П3 PS5", callback_data="pos:P3PS5")
    kb.button(text="⬅️ Назад", callback_data="back")
    kb.adjust(2, 1, 1)
    return kb.as_markup()

def edition_kb(prefix: str, versions: List[str], *, include_all: bool = False, include_new: bool = False):
    kb = InlineKeyboardBuilder()
    for i, version in enumerate(versions):
        kb.button(text=version[:60], callback_data=f"{prefix}:{i}")
    if include_all:
        kb.button(text="Все версии", callback_data=f"{prefix}:all")
    if include_new:
        kb.button(text="➕ Добавить новую версию", callback_data=f"{prefix}:new")
    kb.button(text="⬅️ Назад", callback_data="back")
    kb.adjust(1)
    return kb.as_markup()


def accept_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="✅", callback_data="accept")
    kb.button(text="❌", callback_data="reject")
    kb.button(text="⬅️ Назад", callback_data="back")
    kb.adjust(3)
    return kb.as_markup()

def yes_no_back_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="Да", callback_data="steal_yes")
    kb.button(text="Нет", callback_data="steal_no")
    kb.button(text="⬅️ Назад", callback_data="back")
    kb.adjust(3)
    return kb.as_markup()

def wrap_confirm_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="✅", callback_data="wrap_yes")
    kb.button(text="❌", callback_data="wrap_no")
    kb.button(text="⬅️ Назад", callback_data="back")
    kb.adjust(3)
    return kb.as_markup()


def search_entry_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="🔍 Найти по логину", callback_data="search:find")
    kb.button(text="🗑 Удалить аккаунт", callback_data="search:delete")
    kb.button(text="⬅️ Назад", callback_data="search:cancel")
    kb.adjust(1)
    return kb.as_markup()


def delete_confirm_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="✅ Да, удалить", callback_data="del:yes")
    kb.button(text="❌ Отмена", callback_data="del:cancel")
    kb.adjust(2)
    return kb.as_markup()


def steal_offer_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="✅", callback_data="accept")
    kb.button(text="❌", callback_data="reject")
    kb.button(text="🔢 №", callback_data="steal_jump")
    kb.button(text="⬅️ Назад", callback_data="back")
    kb.adjust(3, 1)
    return kb.as_markup()


def edit_menu_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="✏️ Изменить логин", callback_data="acc:edit:login")
    kb.button(text="🔑 Изменить пароль", callback_data="acc:edit:password")
    kb.button(text="📝 Коды подтверждения", callback_data="acc:codes")
    kb.button(text="⬅️ Назад", callback_data="acc:edit_back")
    kb.adjust(1)
    return kb.as_markup()


def issue_next_code_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="🔄 Другой код", callback_data="issue:nextcode")
    kb.adjust(1)
    return kb.as_markup()


async def safe_edit_text(msg: Message, text: str, reply_markup=None) -> None:
    for attempt in range(10):
        try:
            await msg.edit_text(text, reply_markup=reply_markup)
            return
        except TelegramBadRequest as e:
            err = str(e).lower()
            if "message is not modified" in err:
                try:
                    await msg.edit_reply_markup(reply_markup=reply_markup)
                except TelegramBadRequest:
                    pass
                return
            if "retry" in err or "flood" in err or "too many" in err or "wait" in err:
                await asyncio.sleep(0.35 + attempt * 0.12)
                continue
            raise


async def safe_bot_edit(bot, chat_id: int, message_id: int, text: str, reply_markup=None) -> None:
    for attempt in range(10):
        try:
            await bot.edit_message_text(
                chat_id=chat_id,
                message_id=message_id,
                text=text,
                reply_markup=reply_markup,
            )
            return
        except TelegramBadRequest as e:
            err = str(e).lower()
            if "message is not modified" in err:
                try:
                    await bot.edit_message_reply_markup(
                        chat_id=chat_id,
                        message_id=message_id,
                        reply_markup=reply_markup,
                    )
                except TelegramBadRequest:
                    pass
                return
            if "retry" in err or "flood" in err or "too many" in err or "wait" in err:
                await asyncio.sleep(0.35 + attempt * 0.12)
                continue
            raise


async def try_answer(cq: CallbackQuery, text: str = "", show_alert: bool = False) -> None:
    try:
        await cq.answer(text=text, show_alert=show_alert)
    except TelegramBadRequest:
        pass


def build_issue_success_text(game: str, row_idx: int, pos: str, code: Optional[str], version: str = "") -> str:
    gs = repo.games[game]
    login = repo.get_account_login(game, row_idx)
    password = repo.get_account_password(game, row_idx)
    note = repo.get_account_note(game, row_idx)
    lines = ["✅ Вы выдали этот аккаунт", "", pos, game, f"Версия: {version or 'Standard'}", "", login, password]
    if note:
        lines.extend(["", "⚠️ ПОМЕТКА ДЛЯ МЕНЕДЖЕРА", note])
    if code:
        lines.extend(["", f"Код подтверждения - {code}"])
    return "\n".join(lines)

def render_positions_summary(game: str, row_idx: int) -> str:
    gs = repo.games[game]
    lines: List[str] = []
    cols_p2 = repo._get_cols(game, "П2 PS4/PS5")
    cols_p3p4 = repo._get_cols(game, "П3 PS4")
    cols_p3p5 = repo._get_cols(game, "П3 PS5")

    def add_lines(label: str, cols: List[int]):
        for c in cols:
            val = gs.ws.cell(row_idx, c).value
            has = isinstance(val, str) and val.strip() == "+"
            mark = "✅" if has else "❌"
            lines.append(f"• {label}: {mark}")

    add_lines("П2 PS4/PS5", cols_p2)
    add_lines("П3 PS4", cols_p3p4)
    add_lines("П3 PS5", cols_p3p5)

    body = "\n".join(lines) if lines else "• (нет колонок позиций)"
    return f"Позиции на аккаунте:\n{body}"


def render_codes_section(game: str, row_idx: int) -> str:
    codes = repo.get_codes(game, row_idx)
    if not codes:
        return ""
    lines = ["Неиспользованные коды подтверждения:"]
    for c in codes:
        lines.append(f"• {c}")
    return "\n".join(lines)


def format_account_card(game: str, row_idx: int) -> str:
    gs = repo.games[game]
    login = repo.get_account_login(game, row_idx)
    password = repo.get_account_password(game, row_idx)
    version = repo.get_account_version(game, row_idx) or "Standard"
    note = repo.get_account_note(game, row_idx)
    parts = [game, f"Версия: {version}", "", login, password]
    if note:
        parts.extend(["", "⚠️ ПОМЕТКА ДЛЯ МЕНЕДЖЕРА", note])
    parts.extend(["", render_positions_summary(game, row_idx)])
    codes_block = render_codes_section(game, row_idx)
    if codes_block:
        parts.extend(["", codes_block])
    return "\n".join(parts)


async def refresh_account_message(msg: Message, game: str, row_idx: int, *, show_done: bool = False):
    text = format_account_card(game, row_idx)
    kb = per_cell_buttons(game, row_idx, show_done=show_done)
    await safe_edit_text(msg, text, reply_markup=kb)


def per_cell_buttons(game: str, row_idx: int, *, show_done: bool = False):
    """Кнопки по ячейкам + правка данных + назад; при show_done — завершение сценария «добавить аккаунт»."""
    kb = InlineKeyboardBuilder()
    cols_p2 = repo._get_cols(game, "П2 PS4/PS5")
    cols_p3p4 = repo._get_cols(game, "П3 PS4")
    cols_p3p5 = repo._get_cols(game, "П3 PS5")

    def add_buttons(label: str, cols: List[int]):
        for c in cols:
            has = repo.cell_has_plus(game, row_idx, c)
            text = ("❌ Убрать " + label if has else "➕ Вернуть " + label)
            kb.button(text=text, callback_data=f"acc:cell:{c}")

    add_buttons("П2 PS4/PS5", cols_p2)
    add_buttons("П3 PS4", cols_p3p4)
    add_buttons("П3 PS5", cols_p3p5)

    kb.button(text="⚠️ Пометка", callback_data="acc:note")
    kb.button(text="✏️ Изменить данные", callback_data="acc:edit")
    if show_done:
        kb.button(text="✅ Готово", callback_data="add:done")
    kb.button(text="⬅️ Назад", callback_data="acc:back")
    kb.adjust(1)
    return kb.as_markup()

# ---------- Bot ----------
class MenuText(BaseFilter):
    """Совпадение по подстроке (эмодзи в начале кнопки не мешают)."""

    def __init__(self, substr: str):
        self.substr = substr.lower()

    async def __call__(self, message: Message) -> bool:
        t = (message.text or "").lower()
        return self.substr in t


router = Router()
repo = GameRepo(EXCEL_PATH, MASTER_SHEET)

_issue_locks: Dict[int, asyncio.Lock] = {}


def _issue_lock(chat_id: int) -> asyncio.Lock:
    if chat_id not in _issue_locks:
        _issue_locks[chat_id] = asyncio.Lock()
    return _issue_locks[chat_id]


def admin_check(user_id: int) -> bool:
    return (not ADMIN_IDS) or (user_id in ADMIN_IDS)


STALE_UI_MSG = (
    "Это сообщение устарело. Действуй только с последним экраном сценария или начни заново с меню ниже."
)


async def _active_ui_matches(cq: CallbackQuery, state: FSMContext) -> bool:
    """Один активный message_id на сценарий — старые inline-кнопки не обрабатываем."""
    data = await state.get_data()
    expected = data.get("active_ui_message_id")
    if expected is None:
        return True
    if cq.message.message_id != int(expected):
        await try_answer(cq, STALE_UI_MSG, show_alert=True)
        return False
    return True


# === Excel send/upload ===
@router.message(MenuText("загрузить excel"))
async def send_excel_file(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        await m.answer("Файл Excel не найден. Проверь EXCEL_PATH в .env")
        return

    try:
        doc = FSInputFile(EXCEL_PATH, filename=os.path.basename(EXCEL_PATH))
        await m.answer_document(
            document=doc,
            caption=("Это рабочий Excel-файл. "
                     "Отредактируй и пришли его сюда как документ.\n\n"
                     "Важно: закрой файл в Excel перед отправкой.")
        )
        await state.set_state(ExcelStates.WaitingUpload)
        sent = await m.answer("Жду новый .xlsx ⬇️", reply_markup=back_inline_kb())
        await state.update_data(active_ui_message_id=sent.message_id)
    except Exception as e:
        await m.answer(f"Ошибка при отправке файла: {e}")


@router.message(ExcelStates.WaitingUpload, F.text)
async def excel_waiting_text(m: Message, state: FSMContext):
    if await interrupt_if_main_menu(m, state):
        return
    await m.answer("Пришли файл .xlsx документом или выбери другое действие в меню ниже.")


@router.message(ExcelStates.WaitingUpload, F.document)
async def receive_excel(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return

    doc = m.document
    name = (doc.file_name or "").lower()
    if not name.endswith(".xlsx"):
        await m.answer("Нужен файл .xlsx. Пришли Excel-файл.", reply_markup=back_inline_kb())
        return

    # backup
    try:
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        if os.path.exists(EXCEL_PATH):
            backup = f"{EXCEL_PATH}.bak_{ts}"
            shutil.copy2(EXCEL_PATH, backup)
    except Exception as e:
        await m.answer(f"⚠️ Не удалось сделать бэкап: {e}\nПродолжаю замену.")

    try:
        # direct download
        try:
            await m.bot.download(doc, destination=EXCEL_PATH)
        except Exception:
            file = await m.bot.get_file(doc.file_id)
            await m.bot.download_file(file.file_path, EXCEL_PATH)

        repo.load()
        await m.answer("✅ Файл получен и применён. Excel перезагружен.")
    except Exception as e:
        await m.answer(f"❌ Ошибка при замене файла: {e}")
    finally:
        await state.clear()
        await m.answer("Готово.", reply_markup=main_menu_kb())

# === Start / Main ===
@router.message(CommandStart())
async def start(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    await m.answer(
        "Привет! Панель управления аккаунтами.\n"
        "Выбери действие на клавиатуре ниже.",
        reply_markup=main_menu_kb(),
    )
    await state.clear()

# === Manual save ===
@router.message(F.text == "💾 Сохранить изменения")
async def manual_save_btn(m: Message):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    ok = await async_repo_save()
    await m.answer("✅ Изменения сохранены!" if ok else "❌ Не удалось сохранить.")

@router.message(Command("save"))
async def manual_save_cmd(m: Message):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    ok = await async_repo_save()
    await m.answer("✅ Изменения сохранены!" if ok else "❌ Не удалось сохранить.")


async def interrupt_if_main_menu(m: Message, state: FSMContext) -> bool:
    """Сброс сценария и переход по кнопке нижнего меню."""
    if not m.text or not admin_check(m.from_user.id):
        return False
    t = m.text.lower()
    if "добавить аккаунт" in t:
        await state.clear()
        await state.set_state(AddAccStates.ChoosingGame)
        sent = await m.answer(
            "В какую игру добавляем аккаунт? Напиши название или часть названия.",
            reply_markup=back_inline_kb(),
        )
        await state.update_data(active_ui_message_id=sent.message_id)
        return True
    if "сверить аккаунты" in t:
        await state.clear()
        await state.set_state(AuditStates.WaitingLogins)
        sent = await m.answer("Отправь список почт аккаунтов одним сообщением — по одной почте на строку или через запятую.", reply_markup=back_inline_kb())
        await state.update_data(active_ui_message_id=sent.message_id)
        return True
    if "поиск аккаунта" in t:
        await state.clear()
        await state.set_state(SearchStates.Entry)
        sent = await m.answer("Выбери действие:", reply_markup=search_entry_kb())
        await state.update_data(active_ui_message_id=sent.message_id)
        return True
    if "загрузить excel" in t:
        await state.clear()
        await send_excel_file(m, state)
        return True
    if "выдать игру" in t:
        await state.clear()
        await state.set_state(IssueStates.ChoosingGame)
        sent = await m.answer(
            "Какую выдаём? Напиши название или часть названия.",
            reply_markup=back_inline_kb(),
        )
        await state.update_data(active_ui_message_id=sent.message_id)
        return True
    if "сохранить изменения" in t or (m.text and m.text.strip().startswith("💾")):
        await state.clear()
        ok = await async_repo_save()
        await m.answer("✅ Изменения сохранены!" if ok else "❌ Не удалось сохранить.")
        return True
    return False


# === Issue flow ===
@router.message(MenuText("выдать игру"))
async def issue_begin(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    await state.set_state(IssueStates.ChoosingGame)
    sent = await m.answer(
        "Какую выдаём? Напиши название или часть названия.",
        reply_markup=back_inline_kb(),
    )
    await state.update_data(active_ui_message_id=sent.message_id)

@router.message(IssueStates.ChoosingGame)
async def choose_game_input(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    query = (m.text or "").strip()
    try:
        matches = repo.find_games_fuzzy(query, limit=12, score_cutoff=80)
    except Exception as e:
        await m.answer(f"Ошибка чтения Excel: {e}")
        return
    if not matches:
        await m.answer("Не нашёл совпадений ≥ 80%. Попробуй точнее.", reply_markup=back_inline_kb())
        return
    kb = InlineKeyboardBuilder()
    for name in matches:
        kb.button(text=name[:64], callback_data=f"game:{name}")
    kb.button(text="⬅️ Назад", callback_data="back")
    kb.adjust(1)
    sent = await m.answer("Выбери игру:", reply_markup=kb.as_markup())
    await state.update_data(active_ui_message_id=sent.message_id)

@router.callback_query(F.data.startswith("game:"))
async def select_game(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    game = cq.data.split("game:", 1)[1]
    versions = repo.get_versions(game)
    await state.update_data(game=game, issue_versions=versions)
    await state.set_state(IssueStates.ChoosingVersion)
    await cq.message.edit_text(
        f"Игра: {game}\nСначала выбери версию:",
        reply_markup=edition_kb("issuever", versions, include_all=True),
    )
    await state.update_data(active_ui_message_id=cq.message.message_id)

@router.callback_query(F.data.startswith("issuever:"))
async def select_issue_version(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    data = await state.get_data()
    versions = data.get("issue_versions", [])
    token = cq.data.split(":", 1)[1]
    if token == "all":
        version = None
        label = "Все версии"
    else:
        try:
            version = versions[int(token)]
        except (ValueError, IndexError):
            await cq.answer("Список версий устарел. Начни заново.", show_alert=True)
            return
        label = version
    await cq.answer()
    await state.update_data(issue_version=version, issue_version_label=label)
    game = data["game"]
    await state.set_state(IssueStates.ChoosingPosition)
    await cq.message.edit_text(f"Игра: {game}\nВерсия: {label}\nВыбери позицию:", reply_markup=positions_kb())
    await state.update_data(active_ui_message_id=cq.message.message_id)

@router.callback_query(F.data == "back")
async def go_back(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    st = await state.get_state()
    data = await state.get_data()
    if data.get("active_ui_message_id") is not None:
        if not await _active_ui_matches(cq, state):
            return
    await cq.answer()
    if st == IssueStates.ChoosingVersion.state:
        await state.set_state(IssueStates.ChoosingGame)
        await cq.message.edit_text("Какую выдаём? Напиши название или часть названия.", reply_markup=back_inline_kb())
    elif st == IssueStates.ChoosingPosition.state:
        data = await state.get_data()
        game = data.get("game")
        versions = data.get("issue_versions") or repo.get_versions(game)
        await state.set_state(IssueStates.ChoosingVersion)
        await cq.message.edit_text(f"Игра: {game}\nВыбери версию:", reply_markup=edition_kb("issuever", versions, include_all=True))
    elif st == IssueStates.StealAwaitJump.state:
        await state.set_state(IssueStates.StealOffering)
        await offer_current(cq.message, state)
    elif st == AddAccStates.ChoosingEdition.state:
        await state.set_state(AddAccStates.ChoosingGame)
        await cq.message.edit_text("В какую игру добавляем аккаунт? Напиши название или часть названия.", reply_markup=back_inline_kb())
    elif st in (AddAccStates.WaitingNewEdition.state, AddAccStates.WaitingBulkCreds.state):
        data = await state.get_data()
        game = data.get("add_game")
        versions = repo.get_versions(game)
        await state.update_data(add_versions=versions)
        await state.set_state(AddAccStates.ChoosingEdition)
        await cq.message.edit_text(f"Игра: {game}\nВыбери версию аккаунта или добавь новую:", reply_markup=edition_kb("addver", versions, include_new=True))
    elif st in (AddAccStates.AskRemove.state, AuditStates.WaitingLogins.state):
        await state.clear()
        await cq.message.edit_text("Действие отменено.")
        await cq.message.answer("Главное меню — выбери действие ниже.", reply_markup=main_menu_kb())
    elif st == SearchStates.AwaitNote.state:
        data = await state.get_data()
        game = data.get("game")
        row = data.get("row")
        mode = data.get("mode")
        await state.set_state(SearchStates.Showing)
        await cq.message.edit_text("Изменение пометки отменено.")
        if game and row is not None:
            sent = await cq.message.answer(
                format_account_card(game, row),
                reply_markup=per_cell_buttons(game, row, show_done=(mode == "add")),
            )
            await state.update_data(active_ui_message_id=sent.message_id)
    elif st == SearchStates.AwaitCodes.state:
        data = await state.get_data()
        game = data.get("game")
        row = data.get("row")
        mode = data.get("mode")
        await state.set_state(SearchStates.Showing)
        await cq.message.edit_text("Отменено.")
        if game and row is not None:
            sent = await cq.message.answer(
                format_account_card(game, row),
                reply_markup=per_cell_buttons(game, row, show_done=(mode == "add")),
            )
            await state.update_data(active_ui_message_id=sent.message_id)
    elif st in (
        IssueStates.Offering.state,
        IssueStates.StealDecision.state,
        IssueStates.StealOffering.state,
        IssueStates.StealWrapConfirm.state,
    ):
        game = data.get("game", "—")
        await state.set_state(IssueStates.ChoosingPosition)
        await cq.message.edit_text(f"Игра: {game}\nВыбери позицию:", reply_markup=positions_kb())
    else:
        await cq.message.edit_text("Главное меню. Выбирай действие.")
        await cq.message.answer("Готово.", reply_markup=main_menu_kb())

@router.callback_query(F.data.startswith("pos:"))
async def select_position(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    mapping = {"pos:P2": "П2 PS4/PS5", "pos:P3PS4": "П3 PS4", "pos:P3PS5": "П3 PS5"}
    pos = mapping[cq.data]
    data = await state.get_data()
    game = data["game"]

    selected_version = data.get("issue_version")
    queue: List[Tuple[int, int]] = [(acc.row_idx, col) for acc, col in repo.iter_free_accounts(game, pos, selected_version)]
    await state.update_data(
        position=pos,
        queue=queue,
        ptr=0,
        steal_mode=False,
        current=None,
        steal_idx=0,
    )

    if queue:
        await state.set_state(IssueStates.Offering)
        await offer_current(cq.message, state)
    else:
        await state.set_state(IssueStates.StealDecision)
        await cq.message.edit_text("Закончилась эта позиция. Воровать?", reply_markup=yes_no_back_kb())
        await state.update_data(active_ui_message_id=cq.message.message_id)

async def offer_current(msg: Message, state: FSMContext):
    chat_id = msg.chat.id
    async with _issue_lock(chat_id):
        await _offer_current_impl(msg, state)


async def _offer_current_impl(msg: Message, state: FSMContext):
    data = await state.get_data()
    game = data["game"]
    pos = data.get("position", "")
    queue: List[Tuple[int, int]] = data.get("queue", [])
    ptr: int = data.get("ptr", 0)
    steal_mode: bool = data.get("steal_mode", False)

    if steal_mode:
        gs = repo.games[game]
        selected_version = data.get("issue_version")
        all_rows = [acc.row_idx for acc in gs.accounts if selected_version is None or (repo.get_account_version(game, acc.row_idx) or "Standard").casefold() == selected_version.casefold()]
        n = len(all_rows)
        if n == 0:
            await safe_edit_text(msg, "Больше аккаунтов нет.", reply_markup=back_inline_kb())
            await state.update_data(active_ui_message_id=msg.message_id)
            return

        steal_idx = int(data.get("steal_idx", 0))
        if steal_idx < 0 or steal_idx >= n:
            steal_idx = 0
            await state.update_data(steal_idx=0)

        row = all_rows[steal_idx]
        login = repo.get_account_login(game, row)
        password = repo.get_account_password(game, row)
        await state.update_data(current=(row, None))
        version = repo.get_account_version(game, row) or "Standard"
        note = repo.get_account_note(game, row)
        note_block = f"\n\n⚠️ ПОМЕТКА ДЛЯ МЕНЕДЖЕРА\n{note}" if note else ""
        text = (
            f"Ворование · аккаунт {steal_idx + 1} из {n}\n"
            f"{pos}\n{game}\nВерсия: {version}\n\n{login}\n{password}{note_block}\n\nЭтот аккаунт подойдёт?"
        )
        await safe_edit_text(msg, text, reply_markup=steal_offer_kb())
        await state.update_data(active_ui_message_id=msg.message_id)
        return

    if ptr >= len(queue):
        await state.set_state(IssueStates.StealDecision)
        await safe_edit_text(
            msg,
            "Закончились аккаунты с этой позицией. Воровать?",
            reply_markup=yes_no_back_kb(),
        )
        await state.update_data(active_ui_message_id=msg.message_id)
        return

    row_idx, col_idx = queue[ptr]
    gs = repo.games[game]
    login = repo.get_account_login(game, row_idx)
    password = repo.get_account_password(game, row_idx)

    await state.update_data(current=(row_idx, col_idx))
    version = repo.get_account_version(game, row_idx) or "Standard"
    note = repo.get_account_note(game, row_idx)
    note_block = f"\n\n⚠️ ПОМЕТКА ДЛЯ МЕНЕДЖЕРА\n{note}" if note else ""
    await safe_edit_text(
        msg,
        f"{pos}\n{game}\nВерсия: {version}\n\n{login}\n{password}{note_block}\n\nЭтот аккаунт подойдёт?",
        reply_markup=accept_kb(),
    )
    await state.update_data(active_ui_message_id=msg.message_id)

@router.callback_query(F.data == "reject")
async def reject_offer(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await try_answer(cq, "Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await try_answer(cq)
    await asyncio.sleep(0)
    data = await state.get_data()
    steal_mode: bool = data.get("steal_mode", False)
    if steal_mode:
        game = data["game"]
        gs = repo.games[game]
        selected_version = data.get("issue_version")
        all_rows = [acc.row_idx for acc in gs.accounts if selected_version is None or (repo.get_account_version(game, acc.row_idx) or "Standard").casefold() == selected_version.casefold()]
        n = len(all_rows)
        if n == 0:
            await offer_current(cq.message, state)
            return
        steal_idx = int(data.get("steal_idx", 0))
        old_idx = steal_idx
        steal_idx = (steal_idx + 1) % n
        if steal_idx == 0 and old_idx == n - 1:
            await state.set_state(IssueStates.StealWrapConfirm)
            await state.update_data(steal_idx=0)
            await safe_edit_text(
                cq.message,
                "Прокрутка аккаунтов начинается заново. Продолжаем?",
                reply_markup=wrap_confirm_kb(),
            )
            await state.update_data(active_ui_message_id=cq.message.message_id)
            return
        await state.update_data(steal_idx=steal_idx)
        await offer_current(cq.message, state)
        return

    ptr = data.get("ptr", 0) + 1
    await state.update_data(ptr=ptr)
    await offer_current(cq.message, state)

@router.callback_query(F.data == "accept")
async def accept_offer(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await try_answer(cq, "Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    data = await state.get_data()
    game = data["game"]
    pos = data.get("position", "")
    current = data.get("current")
    steal_mode: bool = data.get("steal_mode", False)

    if not current:
        await try_answer(cq, "Не удалось получить текущий аккаунт.", show_alert=True)
        return

    await try_answer(cq)
    await asyncio.sleep(0)

    if steal_mode:
        row_idx = current[0]
    else:
        row_idx, col_idx = current
        try:
            repo.remove_plus(game, row_idx, col_idx)
        except Exception as e:
            await try_answer(cq, f"Ошибка: {e}", show_alert=True)
            return

    version = repo.get_account_version(game, row_idx) or "Standard"
    login = repo.get_account_login(game, row_idx)
    try:
        repo.append_issue_history(game, row_idx, version, pos, login, cq.from_user.id)
    except Exception as e:
        print(f"[issue history] {e}")
    code = repo.pop_first_code(game, row_idx)
    remaining = repo.get_codes(game, row_idx)
    text = build_issue_success_text(game, row_idx, pos, code, version)
    if remaining:
        await state.set_state(IssueStates.IssuedCode)
        await state.update_data(
            issued_game=game,
            issued_row=row_idx,
            issued_pos=pos,
            issued_version=version,
            issued_steal=steal_mode,
        )
        await safe_edit_text(cq.message, text, reply_markup=issue_next_code_kb())
    else:
        await safe_edit_text(cq.message, text, reply_markup=None)
        await state.clear()

@router.callback_query(F.data == "steal_yes")
async def steal_yes(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    await state.update_data(steal_mode=True, steal_idx=0, current=None)
    await state.set_state(IssueStates.StealOffering)
    await offer_current(cq.message, state)

@router.callback_query(F.data == "steal_jump")
async def steal_jump_ask(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    data = await state.get_data()
    game = data.get("game")
    selected_version = data.get("issue_version")
    n = sum(1 for acc in repo.games[game].accounts if selected_version is None or (repo.get_account_version(game, acc.row_idx) or "Standard").casefold() == selected_version.casefold())
    if n == 0:
        await cq.answer("Нет аккаунтов.", show_alert=True)
        return
    await cq.answer()
    await state.update_data(steal_jump_total=n, active_ui_message_id=cq.message.message_id)
    await state.set_state(IssueStates.StealAwaitJump)
    await cq.message.edit_text(
        f"Напиши номер аккаунта от 1 до {n} (ответом в чат).\n"
        f"Нумерация как строки в таблице Excel (сверху вниз).",
        reply_markup=back_inline_kb(),
    )


@router.message(IssueStates.StealAwaitJump, F.text)
async def steal_jump_apply(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    data = await state.get_data()
    n = int(data.get("steal_jump_total", 0))
    mid = data.get("active_ui_message_id")
    if n <= 0 or not mid:
        await state.clear()
        await m.answer("Сессия сброшена. Начни выдачу снова.", reply_markup=main_menu_kb())
        return
    try:
        num = int((m.text or "").strip())
    except ValueError:
        await m.answer(f"Нужно целое число от 1 до {n}.")
        return
    if num < 1 or num > n:
        await m.answer(f"От 1 до {n}.")
        return
    await state.update_data(steal_idx=num - 1)
    await state.set_state(IssueStates.StealOffering)
    try:
        await m.delete()
    except Exception:
        pass
    async with _issue_lock(m.chat.id):
        data = await state.get_data()
        game = data["game"]
        pos = data.get("position", "")
        gs = repo.games[game]
        selected_version = data.get("issue_version")
        all_rows = [acc.row_idx for acc in gs.accounts if selected_version is None or (repo.get_account_version(game, acc.row_idx) or "Standard").casefold() == selected_version.casefold()]
        n2 = len(all_rows)
        steal_idx = int(data.get("steal_idx", 0))
        if steal_idx < 0 or steal_idx >= n2:
            steal_idx = 0
        row = all_rows[steal_idx]
        login = repo.get_account_login(game, row)
        password = repo.get_account_password(game, row)
        await state.update_data(current=(row, None))
        version = repo.get_account_version(game, row) or "Standard"
        note = repo.get_account_note(game, row)
        note_block = f"\n\n⚠️ ПОМЕТКА ДЛЯ МЕНЕДЖЕРА\n{note}" if note else ""
        text = (
            f"Ворование · аккаунт {steal_idx + 1} из {n2}\n"
            f"{pos}\n{game}\nВерсия: {version}\n\n{login}\n{password}{note_block}\n\nЭтот аккаунт подойдёт?"
        )
        await safe_bot_edit(m.bot, m.chat.id, mid, text, reply_markup=steal_offer_kb())


@router.callback_query(F.data == "issue:nextcode")
async def issue_next_code_handler(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await try_answer(cq, "Доступ ограничен.", show_alert=True)
        return
    if await state.get_state() != IssueStates.IssuedCode.state:
        await try_answer(cq, "Это сообщение устарело.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await try_answer(cq)
    data = await state.get_data()
    game = data.get("issued_game")
    row = data.get("issued_row")
    pos = data.get("issued_pos")
    version = data.get("issued_version", "")
    if not game or row is None:
        await state.clear()
        return
    code = repo.pop_first_code(game, int(row))
    if not code:
        text = build_issue_success_text(game, int(row), pos, None, version)
        await safe_edit_text(cq.message, text, reply_markup=None)
        await state.clear()
        return
    remaining = repo.get_codes(game, int(row))
    text = build_issue_success_text(game, int(row), pos, code, version)
    if remaining:
        await safe_edit_text(cq.message, text, reply_markup=issue_next_code_kb())
    else:
        await safe_edit_text(cq.message, text, reply_markup=None)
        await state.clear()


@router.callback_query(F.data == "steal_no")
async def steal_no(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    data = await state.get_data()
    game = data.get("game", "—")
    await state.set_state(IssueStates.ChoosingPosition)
    await cq.message.edit_text(f"Игра: {game}\nВыбери позицию:", reply_markup=positions_kb())

@router.callback_query(F.data == "wrap_yes")
async def wrap_yes(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    await state.update_data(steal_idx=0, current=None)
    await state.set_state(IssueStates.StealOffering)
    await offer_current(cq.message, state)

@router.callback_query(F.data == "wrap_no")
async def wrap_no(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    data = await state.get_data()
    game = data.get("game", "—")
    await state.set_state(IssueStates.ChoosingPosition)
    await cq.message.edit_text(f"Игра: {game}\nВыбери позицию:", reply_markup=positions_kb())

@router.message(Command("reload"))
async def reload_cmd(m: Message):
    if not admin_check(m.from_user.id):
        await m.reply("Доступ ограничен.")
        return
    try:
        repo.load()
        await m.reply("Excel перезагружен.")
    except Exception as e:
        await m.reply(f"Ошибка: {e}")

@router.message(Command("help"))
async def help_cmd(m: Message):
    if not admin_check(m.from_user.id):
        await m.reply("Доступ ограничен.")
        return
    await m.reply(
        "Команды:\n"
        "• /reload — перечитать Excel\n"
        "• /save — сохранить Excel\n\n"
        "Клавиатура: выдача, загрузка Excel, поиск и добавление аккаунта, сохранение.",
        reply_markup=main_menu_kb(),
    )

# === Поиск аккаунта + Редактирование ===
@router.message(MenuText("поиск аккаунта"))
async def search_start(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    await state.set_state(SearchStates.Entry)
    sent = await m.answer("Выбери действие:", reply_markup=search_entry_kb())
    await state.update_data(active_ui_message_id=sent.message_id)


@router.callback_query(F.data == "search:find")
async def search_find_cb(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    await state.set_state(SearchStates.WaitingLogin)
    await cq.message.edit_text("Пришли логин (email) аккаунта.", reply_markup=back_inline_kb())


@router.callback_query(F.data == "search:delete")
async def search_delete_cb(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    await state.set_state(SearchStates.AwaitDeleteLogin)
    await cq.message.edit_text(
        "Пришли логин (email) аккаунта, который нужно удалить из таблицы.",
        reply_markup=back_inline_kb(),
    )


@router.callback_query(F.data == "search:cancel")
async def search_cancel_cb(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    await state.clear()
    await cq.message.edit_text("Главное меню.", reply_markup=None)
    await cq.message.answer("Готово.", reply_markup=main_menu_kb())


@router.message(SearchStates.Entry)
async def search_entry_stray_text(m: Message, state: FSMContext):
    if await interrupt_if_main_menu(m, state):
        return
    await m.answer("Нажми кнопку под предыдущим сообщением или выбери пункт в меню ниже.")


@router.message(SearchStates.WaitingLogin)
async def search_by_login(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return

    login = (m.text or "").strip()
    found = repo.find_account_by_login(login)
    if not found:
        await m.answer("Не нашёл такой логин. Проверь написание.", reply_markup=back_inline_kb())
        return

    game, acc = found

    await state.update_data(mode="search", game=game, row=acc.row_idx)
    await state.set_state(SearchStates.Showing)
    sent = await m.answer(
        format_account_card(game, acc.row_idx),
        reply_markup=per_cell_buttons(game, acc.row_idx),
    )
    await state.update_data(active_ui_message_id=sent.message_id)


@router.message(SearchStates.AwaitDeleteLogin)
async def delete_by_login(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    login = (m.text or "").strip()
    found = repo.find_account_by_login(login)
    if not found:
        await m.answer("Не нашёл такой логин. Проверь написание.", reply_markup=back_inline_kb())
        return
    game, acc = found
    await state.update_data(del_game=game, del_row=acc.row_idx, del_login=acc.login)
    sent = await m.answer(
        f"Удалить аккаунт «{acc.login}» из игры «{game}»?\n"
        f"Строка будет удалена из листа Excel.",
        reply_markup=delete_confirm_kb(),
    )
    await state.update_data(active_ui_message_id=sent.message_id)


@router.callback_query(F.data == "del:yes")
async def delete_confirm_yes(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    data = await state.get_data()
    game = data.get("del_game")
    row = data.get("del_row")
    if not game or row is None:
        await cq.answer("Данные устарели.", show_alert=True)
        return
    await cq.answer()
    repo.delete_account_row(game, int(row))
    await state.clear()
    await cq.message.edit_text("✅ Аккаунт удалён из списка этой игры.")
    await cq.message.answer("Готово.", reply_markup=main_menu_kb())


@router.callback_query(F.data == "del:cancel")
async def delete_confirm_no(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    await state.set_state(SearchStates.Entry)
    await cq.message.edit_text("Выбери действие:", reply_markup=search_entry_kb())


@router.callback_query(F.data == "acc:back")
async def search_back(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    data = await state.get_data()
    mode = data.get("mode")
    await state.clear()
    if mode == "add":
        await cq.message.edit_text("✅ Аккаунт добавлен.\n\nГлавное меню — выбери действие ниже.", reply_markup=None)
    else:
        await cq.message.edit_text("Главное меню — выбери действие ниже.", reply_markup=None)
    await cq.message.answer("Готово.", reply_markup=main_menu_kb())

@router.callback_query(F.data.startswith("acc:cell:"))
async def toggle_specific_cell(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return

    data = await state.get_data()
    if data.get("mode") not in ("search", "add"):
        await cq.answer("Контекст потерян. Начни снова: поиск или добавление аккаунта.", show_alert=True)
        return

    game = data["game"]
    row = data["row"]
    try:
        col = int(cq.data.split("acc:cell:", 1)[1])
    except ValueError:
        await cq.answer("Некорректная колонка.", show_alert=True)
        return

    has = repo.cell_has_plus(game, row, col)
    repo.set_cell_plus(game, row, col, not has)

    show_done = data.get("mode") == "add"
    await refresh_account_message(cq.message, game, row, show_done=show_done)
    await cq.answer("Убрано" if has else "Добавлено")

@router.callback_query(F.data == "acc:edit")
async def edit_entrypoint(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    await state.set_state(SearchStates.EditMenu)
    await cq.message.edit_reply_markup(reply_markup=edit_menu_kb())

@router.callback_query(F.data == "acc:edit_back")
async def edit_back_to_show(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    data = await state.get_data()
    game = data["game"]
    row = data["row"]
    show_done = data.get("mode") == "add"
    await state.set_state(SearchStates.Showing)
    await cq.message.edit_text(
        format_account_card(game, row),
        reply_markup=per_cell_buttons(game, row, show_done=show_done),
    )

@router.callback_query(F.data == "acc:edit:login")
async def edit_login_ask(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    await state.set_state(SearchStates.AwaitLogin)
    sent = await cq.message.answer(
        "Отправь новый логин для этого аккаунта.\n\n(Чтобы отменить — нажми «⬅️ Назад»)",
        reply_markup=back_inline_kb(),
    )
    await state.update_data(active_ui_message_id=sent.message_id)

@router.callback_query(F.data == "acc:edit:password")
async def edit_password_ask(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    await state.set_state(SearchStates.AwaitPassword)
    sent = await cq.message.answer(
        "Отправь новый пароль для этого аккаунта.\n\n(Чтобы отменить — нажми «⬅️ Назад»)",
        reply_markup=back_inline_kb(),
    )
    await state.update_data(active_ui_message_id=sent.message_id)


@router.callback_query(F.data == "acc:note")
async def acc_note_prompt(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    data = await state.get_data()
    game, row = data.get("game"), data.get("row")
    current = repo.get_account_note(game, row) if game and row else ""
    await state.set_state(SearchStates.AwaitNote)
    prompt = "Отправь пометку для менеджеров. Она будет видна при поиске и выдаче аккаунта.\n\nЧтобы удалить текущую пометку, отправь: -"
    if current:
        prompt += f"\n\nТекущая пометка: {current}"
    sent = await cq.message.answer(prompt, reply_markup=back_inline_kb())
    await state.update_data(active_ui_message_id=sent.message_id)


@router.message(SearchStates.AwaitNote)
async def acc_note_save(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    data = await state.get_data()
    game, row = data.get("game"), data.get("row")
    note = (m.text or "").strip()
    if len(note) > 500:
        await m.answer("Пометка слишком длинная. Максимум 500 символов.", reply_markup=back_inline_kb())
        return
    repo.set_account_note(game, row, "" if note in ("-", "—", "/clear") else note)
    show_done = data.get("mode") == "add"
    await state.set_state(SearchStates.Showing)
    await m.answer("✅ Пометка удалена." if note in ("-", "—", "/clear") else "✅ Пометка сохранена.")
    sent = await m.answer(format_account_card(game, row), reply_markup=per_cell_buttons(game, row, show_done=show_done))
    await state.update_data(active_ui_message_id=sent.message_id)


@router.callback_query(F.data == "acc:codes")
async def acc_codes_prompt(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    data = await state.get_data()
    if data.get("mode") not in ("search", "add"):
        await cq.answer("Контекст потерян.", show_alert=True)
        return
    await cq.answer()
    await state.set_state(SearchStates.AwaitCodes)
    sent = await cq.message.answer(
        "Пришли коды: одна строка — один код.\n"
        "Текущий список в Excel будет заменён на присланный.",
        reply_markup=back_inline_kb(),
    )
    await state.update_data(active_ui_message_id=sent.message_id)


@router.message(SearchStates.AwaitCodes)
async def acc_codes_save(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    data = await state.get_data()
    game = data["game"]
    row = data["row"]
    lines = [ln.strip() for ln in (m.text or "").splitlines() if ln.strip()]
    repo.set_codes(game, row, lines)
    show_done = data.get("mode") == "add"
    await state.set_state(SearchStates.Showing)
    await m.answer("✅ Коды сохранены в таблице.")
    sent = await m.answer(
        format_account_card(game, row),
        reply_markup=per_cell_buttons(game, row, show_done=show_done),
    )
    await state.update_data(active_ui_message_id=sent.message_id)


@router.message(SearchStates.AwaitLogin)
async def edit_login_set(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    new_login = (m.text or "").strip()
    data = await state.get_data()
    game = data["game"]
    row = data["row"]
    repo.set_login(game, row, new_login)

    show_done = data.get("mode") == "add"
    await state.set_state(SearchStates.Showing)
    await m.answer("✅ Логин обновлён.")
    sent = await m.answer(
        format_account_card(game, row),
        reply_markup=per_cell_buttons(game, row, show_done=show_done),
    )
    await state.update_data(active_ui_message_id=sent.message_id)

@router.message(SearchStates.AwaitPassword)
async def edit_password_set(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    new_password = (m.text or "").strip()
    data = await state.get_data()
    game = data["game"]
    row = data["row"]
    repo.set_password(game, row, new_password)

    show_done = data.get("mode") == "add"
    await state.set_state(SearchStates.Showing)
    await m.answer("✅ Пароль обновлён.")
    sent = await m.answer(
        format_account_card(game, row),
        reply_markup=per_cell_buttons(game, row, show_done=show_done),
    )
    await state.update_data(active_ui_message_id=sent.message_id)

# === Добавить аккаунт ===
@router.message(MenuText("добавить аккаунт"))
async def add_acc_begin(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    await state.set_state(AddAccStates.ChoosingGame)
    sent = await m.answer(
        "В какую игру добавляем аккаунт? Напиши название или часть названия.",
        reply_markup=back_inline_kb(),
    )
    await state.update_data(active_ui_message_id=sent.message_id)

@router.message(AddAccStates.ChoosingGame)
async def add_acc_choose_game(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    query = (m.text or "").strip()
    try:
        matches = repo.find_games_fuzzy(query, limit=12, score_cutoff=80)
    except Exception as e:
        await m.answer(f"Ошибка чтения Excel: {e}")
        return
    if not matches:
        await m.answer("Не нашёл совпадений ≥ 80%. Попробуй точнее.", reply_markup=back_inline_kb())
        return
    kb = InlineKeyboardBuilder()
    for name in matches:
        kb.button(text=name[:64], callback_data=f"addgame:{name}")
    kb.button(text="⬅️ Назад", callback_data="back")
    kb.adjust(1)
    sent = await m.answer("Выбери игру:", reply_markup=kb.as_markup())
    await state.update_data(active_ui_message_id=sent.message_id)

@router.callback_query(F.data.startswith("addgame:"))
async def add_acc_selected_game(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    await cq.answer()
    game = cq.data.split("addgame:", 1)[1]
    versions = repo.get_versions(game)
    await state.update_data(add_game=game, add_versions=versions)
    await state.set_state(AddAccStates.ChoosingEdition)
    await cq.message.edit_text(
        f"Игра: {game}\nВыбери версию аккаунта или добавь новую:",
        reply_markup=edition_kb("addver", versions, include_new=True),
    )
    await state.update_data(active_ui_message_id=cq.message.message_id)

@router.callback_query(F.data.startswith("addver:"))
async def add_acc_choose_edition(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    data = await state.get_data()
    token = cq.data.split(":", 1)[1]
    if token == "new":
        await cq.answer()
        await state.set_state(AddAccStates.WaitingNewEdition)
        await cq.message.edit_text("Напиши название версии, например Standard, Ultimate, Deluxe или Gold:", reply_markup=back_inline_kb())
        await state.update_data(active_ui_message_id=cq.message.message_id)
        return
    versions = data.get("add_versions", [])
    try:
        version = versions[int(token)]
    except (ValueError, IndexError):
        await cq.answer("Список версий устарел. Начни добавление заново.", show_alert=True)
        return
    if version == "Standard":
        version = "Standard"
    await cq.answer()
    await state.update_data(add_version=version)
    await state.set_state(AddAccStates.WaitingBulkCreds)
    game = data.get("add_game")
    prompt = (f"Игра: {game}\nВерсия: {version or 'Standard'}\n\n"
              "Отправь данные аккаунта или сразу нескольких аккаунтов одним сообщением либо .txt-файлом.\n"
              "Поддерживаются обычные пары логин/пароль и сообщения поставщика с заголовками «Аккаунт 1 из N».\n"
              "Для поставщицкого формата бот сохранит только почту, пароль аккаунта и коды подтверждения; пароль почты и длинный ключ восстановления пропустит.")
    await cq.message.edit_text(prompt, reply_markup=back_inline_kb())
    await state.update_data(active_ui_message_id=cq.message.message_id)

@router.message(AddAccStates.WaitingNewEdition)
async def add_acc_new_edition_text(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    version = repo.canonicalize_version_name((await state.get_data()).get("add_game", ""), (m.text or "").strip())
    if not version or len(version) > 64:
        await m.answer("Название версии должно содержать от 1 до 64 символов.", reply_markup=back_inline_kb())
        return
    await state.update_data(add_version=version)
    await state.set_state(AddAccStates.WaitingBulkCreds)
    data = await state.get_data()
    sent = await m.answer(
        f"Игра: {data.get('add_game')}\nВерсия: {version}\n\n"
        "Отправь данные одного или нескольких аккаунтов одним сообщением либо .txt-файлом.\n"
        "Можно отправлять обычный список или сообщение поставщика с блоками «Аккаунт N из M».",
        reply_markup=back_inline_kb(),
    )
    await state.update_data(active_ui_message_id=sent.message_id)

def _clean_supplier_line(line: str) -> str:
    line = line.replace("\\", " ")
    line = re.sub(r"\*\*", "", line)
    return line.strip()


def _is_recovery_key(token: str) -> bool:
    compact = re.sub(r"[^A-Za-z0-9]", "", token)
    return len(compact) >= 70 and bool(re.fullmatch(r"[A-Z0-9]+", compact))


def _supplier_account_from_line(line: str):
    """Extract (email, account password, recovery codes) from supplier credential rows."""
    email_match = re.search(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", line)
    if not email_match:
        return None
    email = email_match.group(0).strip()
    before = line[:email_match.start()]
    after = line[email_match.end():]
    # Supplier format A: email ACCOUNT_PASSWORD | MAIL_PASSWORD LONG_RECOVERY_KEY CODES...
    if "|" in after:
        left_tail, right = after.split("|", 1)
        left_tokens = left_tail.split()
        right_tokens = right.split()
        password = left_tokens[0] if left_tokens else (right_tokens[0] if right_tokens else "")
        # Ignore mailbox password and the long recovery key; retain only 6-char recovery codes after key.
        key_index = next((i for i, t in enumerate(right_tokens) if _is_recovery_key(t)), None)
        codes = right_tokens[key_index + 1:] if key_index is not None else []
    else:
        # Supplier format B: email MAIL_PASSWORD ACCOUNT_PASSWORD LONG_RECOVERY_KEY CODES...
        tokens = after.split()
        if not tokens:
            return None
        known_account_passwords = {"joy.fp.10!", "joy.fp.10"}
        known_idx = next((i for i, t in enumerate(tokens[:4]) if t.casefold() in known_account_passwords), None)
        if known_idx is not None:
            password = tokens[known_idx]
        else:
            # If supplier changes both passwords, per request take the first one after email.
            password = tokens[0]
        key_index = next((i for i, t in enumerate(tokens) if _is_recovery_key(t)), None)
        codes = tokens[key_index + 1:] if key_index is not None else []
    codes = [re.sub(r"[^A-Za-z0-9]", "", c) for c in codes]
    codes = [c for c in codes if 5 <= len(c) <= 12 and not _is_recovery_key(c)]
    return email, password.strip(), codes


def parse_bulk_accounts(raw_text: str):
    """Parse plain credential lists and supplier order messages into account tuples."""
    raw_text = (raw_text or "").replace("\r", "")
    lines = [_clean_supplier_line(x) for x in raw_text.split("\n")]
    lines = [x for x in lines if x]
    if not lines:
        return [], "Список пустой."

    # Supplier orders contain one email-bearing line per account; all surrounding
    # order metadata (prices, game name, account counters, seller name) is ignored.
    supplier_rows = []
    for line in lines:
        parsed = _supplier_account_from_line(line)
        if parsed:
            supplier_rows.append(parsed)
    if supplier_rows:
        return supplier_rows, ""

    # Compact format: one account per line: login | password | code1;code2
    compact = [line for line in lines if line != "---"]
    if compact and all("|" in line for line in compact):
        accounts = []
        for line in compact:
            parts = [x.strip() for x in line.split("|")]
            if len(parts) < 2 or not parts[0] or not parts[1]:
                return [], f"Не удалось разобрать строку: {line}"
            codes = [x.strip() for x in re.split(r"[,;\s]+", " ".join(parts[2:])) if x.strip()]
            accounts.append((parts[0], parts[1], codes))
        return accounts, ""

    blocks, current = [], []
    for line in raw_text.split("\n"):
        line = line.strip()
        if not line or line == "---":
            if current:
                blocks.append(current); current = []
        else:
            current.append(line)
    if current:
        blocks.append(current)
    if len(blocks) == 1 and len(blocks[0]) > 2 and len(blocks[0]) % 2 == 0:
        vals = blocks[0]
        # Treat as pairs only when every even-positioned value looks like a login/email.
        if "@" in vals[0] and all("@" in vals[i] for i in range(0, len(vals), 2)):
            return [(vals[i], vals[i+1], []) for i in range(0, len(vals), 2)], ""
        return [(vals[0], vals[1], vals[2:])], ""
    accounts = []
    for i, block in enumerate(blocks, 1):
        if len(block) < 2:
            return [], f"В блоке №{i} не хватает пароля."
        accounts.append((block[0], block[1], block[2:]))
    return accounts, ""


async def _add_accounts_from_payload(m: Message, state: FSMContext, raw_text: str):
    accounts, error = parse_bulk_accounts(raw_text)
    if error:
        await m.answer(f"❌ {error}\nПроверь формат и отправь данные ещё раз.", reply_markup=back_inline_kb())
        return
    if not accounts:
        await m.answer("Не нашёл аккаунтов в сообщении или файле.", reply_markup=back_inline_kb())
        return
    logins = [a[0].casefold() for a in accounts]
    duplicates = sorted({x for x in logins if logins.count(x) > 1})
    if duplicates:
        await m.answer("В списке есть повторяющиеся логины: " + ", ".join(duplicates[:10]) + ". Удали дубли и отправь заново.", reply_markup=back_inline_kb())
        return
    data = await state.get_data()
    game = data.get("add_game")
    version = data.get("add_version", "Standard")
    added = []
    try:
        for login, password, codes in accounts:
            row = repo.add_account_next_row(game, login, password, codes=codes or None, version=version)
            added.append((row, login))
    except Exception as e:
        await m.answer(f"⚠️ Ошибка после добавления {len(added)} из {len(accounts)} аккаунтов: {e}\nПроверь Excel перед повторной отправкой.")
        return
    await state.clear()
    await m.answer(
        f"✅ Добавлено аккаунтов: {len(added)}\nИгра: {game}\nВерсия: {version}\n\n"
        "Все стандартные позиции отмечены доступными (+). При необходимости их можно изменить через поиск.",
        reply_markup=main_menu_kb(),
    )


@router.message(AddAccStates.WaitingBulkCreds, F.text)
async def add_acc_receive_bulk(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    await _add_accounts_from_payload(m, state, m.text or "")


@router.message(AddAccStates.WaitingBulkCreds, F.document)
async def add_acc_receive_bulk_file(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    doc = m.document
    if not (doc.file_name or "").lower().endswith((".txt", ".log")):
        await m.answer("Для добавления аккаунтов пришли текстовое сообщение или файл .txt.", reply_markup=back_inline_kb())
        return
    if doc.file_size and doc.file_size > 20 * 1024 * 1024:
        await m.answer("Файл слишком большой. Максимальный размер — 20 МБ.", reply_markup=back_inline_kb())
        return
    buf = BytesIO()
    await m.bot.download(doc, destination=buf)
    try:
        raw_text = buf.getvalue().decode("utf-8-sig")
    except UnicodeDecodeError:
        raw_text = buf.getvalue().decode("cp1251", errors="replace")
    await _add_accounts_from_payload(m, state, raw_text)


@router.callback_query(F.data == "add:done")
async def add_acc_done(cq: CallbackQuery, state: FSMContext):
    if not admin_check(cq.from_user.id):
        await cq.answer("Доступ ограничен.", show_alert=True)
        return
    if not await _active_ui_matches(cq, state):
        return
    data = await state.get_data()
    if data.get("mode") != "add":
        await cq.answer("Сначала добавь аккаунт.", show_alert=True)
        return
    await cq.answer()
    await state.clear()
    await cq.message.edit_text("✅ Аккаунт добавлен.\n\nГлавное меню — выбери действие ниже.", reply_markup=None)
    await cq.message.answer("Готово.", reply_markup=main_menu_kb())

# === Сверка аккаунтов по списку логинов ===
@router.message(MenuText("сверить аккаунты"))
async def audit_begin(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    await state.clear()
    await state.set_state(AuditStates.WaitingLogins)
    sent = await m.answer(
        "Отправь список почт одним сообщением или прикрепи .txt-файл с любым количеством адресов.\n"
        "Бот выполнит одну общую сверку и пришлёт один Excel-файл с почтой, игрой и версией.",
        reply_markup=back_inline_kb(),
    )
    await state.update_data(active_ui_message_id=sent.message_id)


def _extract_emails(raw: str) -> List[str]:
    found = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", raw or "")
    unique, seen = [], set()
    for email in found:
        key = email.casefold()
        if key not in seen:
            unique.append(email)
            seen.add(key)
    return unique


async def _run_account_audit(m: Message, state: FSMContext, raw: str):
    unique = _extract_emails(raw)
    if not unique:
        await m.answer("Не нашёл почт. Отправь файл или список с адресами электронной почты.", reply_markup=back_inline_kb())
        return
    wb = __import__("openpyxl").Workbook()
    ws = wb.active
    ws.title = "Сверка аккаунтов"
    ws.append(["Почта", "Игра и версия"])
    found_count = 0
    missing = []
    duplicate_matches = 0
    for login in unique:
        matches = repo.find_accounts_by_login(login)
        if matches:
            found_count += 1
            if len(matches) > 1:
                duplicate_matches += 1
            for game, acc in matches:
                version = repo.get_account_version(game, acc.row_idx) or "Standard"
                ws.append([login, f"{game} — {version}"])
        else:
            missing.append(login)
            ws.append([login, "НЕ НАЙДЕН"])
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 58
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    await m.answer_document(
        BufferedInputFile(output.getvalue(), filename=f"sverka_akkauntov_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"),
        caption=(f"Сверка завершена.\nВсего почт: {len(unique)}\nНайдено: {found_count}\nНе найдено: {len(missing)}\n"
                 f"Повторные совпадения в разных играх: {duplicate_matches}"),
    )
    if missing:
        preview = "\n".join(f"• {x}" for x in missing[:40])
        tail = f"\n… и ещё {len(missing)-40}" if len(missing) > 40 else ""
        await m.answer(f"❌ Не найдены в базе ({len(missing)}):\n{preview}{tail}")
    else:
        await m.answer("✅ Все переданные аккаунты найдены в базе.")
    await state.clear()
    await m.answer("Готово.", reply_markup=main_menu_kb())


@router.message(AuditStates.WaitingLogins, F.text)
async def audit_receive_logins(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    await _run_account_audit(m, state, m.text or "")


@router.message(AuditStates.WaitingLogins, F.document)
async def audit_receive_txt(m: Message, state: FSMContext):
    if not admin_check(m.from_user.id):
        await m.answer("Доступ ограничен.")
        return
    if await interrupt_if_main_menu(m, state):
        return
    doc = m.document
    if not (doc.file_name or "").lower().endswith((".txt", ".csv", ".log")):
        await m.answer("Пришли список почт в файле .txt (подойдут также .csv и .log).", reply_markup=back_inline_kb())
        return
    if doc.file_size and doc.file_size > 20 * 1024 * 1024:
        await m.answer("Файл слишком большой. Максимальный размер — 20 МБ.", reply_markup=back_inline_kb())
        return
    buf = BytesIO()
    await m.bot.download(doc, destination=buf)
    try:
        raw = buf.getvalue().decode("utf-8-sig")
    except UnicodeDecodeError:
        raw = buf.getvalue().decode("cp1251", errors="replace")
    await _run_account_audit(m, state, raw)


# === Resilient runner with autosave ===
autosave_task = None
_repo_save_lock: Optional[asyncio.Lock] = None


def _get_repo_save_lock() -> asyncio.Lock:
    global _repo_save_lock
    if _repo_save_lock is None:
        _repo_save_lock = asyncio.Lock()
    return _repo_save_lock


def _sync_repo_save() -> bool:
    """Синхронное сохранение книги (вызывать только из to_thread / shutdown)."""
    if repo and repo.wb:
        return repo.save()
    return False


async def async_repo_save() -> bool:
    """
    Сохранение Excel без блокировки event loop: polling и ответы бота не зависают.
    Ошибки в логе про Telegram (Bad Gateway, timeout) чаще из‑за сети/серверов API;
    длинный синхронный save раньше мог усугублять таймауты.
    """
    async with _get_repo_save_lock():
        return await asyncio.to_thread(_sync_repo_save)


async def autosave_loop():
    while True:
        try:
            await asyncio.sleep(AUTOSAVE_SECONDS)
            if repo and repo.wb:
                ok = await async_repo_save()
                if ok:
                    print(f"[OK] autosave done ({datetime.now().strftime('%H:%M:%S')})")
        except asyncio.CancelledError:
            break
        except Exception as e:
            print(f"[autosave] пропуск: {e}")
            await asyncio.sleep(5)

async def run_bot():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN не задан в .env")
    if not EXCEL_PATH:
        raise RuntimeError("EXCEL_PATH не задан в .env")

    global repo, autosave_task
    repo = GameRepo(EXCEL_PATH, MASTER_SHEET)
    repo.load()
    # Persist automatic edition/note migration before Telegram polling begins.
    try:
        if not await async_repo_save():
            print("[migration] Excel migration loaded in memory but initial save failed")
    except Exception as e:
        print(f"[migration] initial save failed: {e}")

    bot = Bot(BOT_TOKEN, timeout=ClientTimeout(total=300))
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)

    # стартуем автосейв
    autosave_task = asyncio.create_task(autosave_loop())

    # autorestart polling
    backoff = 1
    try:
        while True:
            try:
                await dp.start_polling(bot)
            except Exception as e:
                # попытка сохранить при креше
                try:
                    await async_repo_save()
                except Exception:
                    pass
                try:
                    print(f"[Runner] Polling crashed: {e}. Restarting in {backoff}s...")
                except Exception:
                    pass
                await asyncio.sleep(backoff)
                backoff = min(backoff * 2, 30)
    finally:
        # финальный сейв при завершении
        try:
            if autosave_task:
                autosave_task.cancel()
                with contextlib.suppress(Exception):
                    await autosave_task
        except Exception:
            pass
        try:
            await async_repo_save()
        except Exception:
            pass

if __name__ == "__main__":
    try:
        asyncio.run(run_bot())
    except KeyboardInterrupt:
        # финальный сейв уже выполнится в finally run_bot
        pass
