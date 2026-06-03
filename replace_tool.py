#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import re
import json
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


# ---------------------------------------------------------------------------
# 置換ロジック（1パス同時置換）
# ---------------------------------------------------------------------------

def _make_pat(old: str, fullhalf_sensitive: bool) -> str:
    """old 文字列を正規表現パターン文字列に変換する。"""
    if fullhalf_sensitive:
        return re.escape(old)
    parts = []
    for ch in old:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:        # 全角 ASCII → 半角も許容
            half = chr(code - 0xFEE0)
            parts.append(f'(?:{re.escape(ch)}|{re.escape(half)})')
        elif 0x21 <= code <= 0x7E:          # 半角 ASCII → 全角も許容
            full = chr(code + 0xFEE0)
            parts.append(f'(?:{re.escape(ch)}|{re.escape(full)})')
        elif ch == ' ':
            parts.append(r'(?:　| )')
        elif ch == '　':
            parts.append(r'(?:　| )')
        else:
            parts.append(re.escape(ch))
    return ''.join(parts)


def build_combined_replacer(pairs_raw: list[tuple[str, str]],
                             case_sensitive: bool,
                             fullhalf_sensitive: bool):
    """
    全ペアを 1 本の正規表現で同時処理する冪等な置換関数を返す。

    【問題と解決策】
        "MS"→"bMS社" のような置換では、出力値 "bMS社" の中に検索パターン "MS"
        が含まれるため、複数回実行すると bMS社→bbMS社社→… と無限増殖する。

        解決: 全ての置換後の値（new）を「ガードパターン」として正規表現の
        先頭に登録し、最高優先度でマッチさせる。ガードがマッチした位置は
        そのまま返す（identity）ので、内部の短いパターンが再ヒットしない。
        → 何度実行しても同じ結果（冪等性）が保証される。
    """
    # (old文字列, patternStr, new文字列) のタプルで保持
    entries = []
    for old, new in pairs_raw:
        if old:
            entries.append((old, _make_pat(old, fullhalf_sensitive), new))

    if not entries:
        return None

    # 長い old を優先（短いパターンが長いパターンの一部にマッチするのを防止）
    # 例: "MSFT" > "MS" の順にすることで "MSFT" が "MS" より先にマッチする
    entries.sort(key=lambda e: len(e[0]), reverse=True)

    flags = re.UNICODE | (0 if case_sensitive else re.IGNORECASE)

    # --- ガードパターン: 置換後の値を長い順に並べて最優先登録 ---
    # 既に置換済みのテキスト（例: "bMS社"）がこのパターンにマッチしたら
    # そのまま返し、内部の短いパターン（"MS"）への再ヒットを防ぐ。
    guard_values = sorted(
        set(new for _, new in pairs_raw if new),
        key=len, reverse=True,  # 長いものを優先（部分マッチ防止）
    )
    guard_pats = [re.escape(v) for v in guard_values]

    if not case_sensitive:
        guard_set = {v.lower() for v in guard_values}
    else:
        guard_set = set(guard_values)

    # 各検索パターンを個別コンパイル（マッチ→置換値の逆引き用）
    individual = [(re.compile(f'(?:{pat})', flags), new) for _, pat, new in entries]

    # 組み合わせ正規表現: ガード優先 → 検索パターン（長い順）
    all_pats = guard_pats + [pat for _, pat, _ in entries]
    combined = re.compile('|'.join(f'(?:{p})' for p in all_pats), flags)

    def replacer(text: str) -> str:
        def sub(m: re.Match) -> str:
            s = m.group(0)
            # ガード判定: 置換後の値にマッチしたらそのまま返す
            s_key = s.lower() if not case_sensitive else s
            if s_key in guard_set:
                return s
            # 検索パターンにマッチした場合は置換値を返す
            for irx, new in individual:
                if irx.fullmatch(s):
                    return new
            return s
        return combined.sub(sub, text)

    return replacer


# --- Office ファイル処理 ---------------------------------------------------

def _process_text_frame(tf, replacer) -> int:
    count = 0
    for para in tf.paragraphs:
        for run in para.runs:
            new = replacer(run.text)
            if new != run.text:
                run.text = new
                count += 1
    return count


def replace_in_pptx(filepath: str, replacer) -> int:
    from pptx import Presentation
    prs   = Presentation(filepath)
    total = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                total += _process_text_frame(shape.text_frame, replacer)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        total += _process_text_frame(cell.text_frame, replacer)
    if total:
        prs.save(filepath)
    return total


def replace_in_xlsx(filepath: str, replacer) -> int:
    import openpyxl
    wb    = openpyxl.load_workbook(filepath)
    total = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    new = replacer(cell.value)
                    if new != cell.value:
                        cell.value = new
                        total += 1
    if total:
        wb.save(filepath)
    return total


def replace_in_docx(filepath: str, replacer) -> int:
    from docx import Document

    def proc_paras(paragraphs):
        n = 0
        for para in paragraphs:
            for run in para.runs:
                new = replacer(run.text)
                if new != run.text:
                    run.text = new
                    n += 1
        return n

    doc   = Document(filepath)
    total = proc_paras(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total += proc_paras(cell.paragraphs)
    if total:
        doc.save(filepath)
    return total


HANDLERS = {
    ".pptx": replace_in_pptx,
    ".xlsx": replace_in_xlsx,
    ".docx": replace_in_docx,
}


# ---------------------------------------------------------------------------
# JSON コメント除去（// スタイル）
# ---------------------------------------------------------------------------

def _strip_json_comments(text: str) -> str:
    """
    // 行コメントを除去して純粋な JSON 文字列を返す。
    文字列リテラル内の // は除去しない。
    """
    result = []
    in_str  = False
    i = 0
    n = len(text)
    while i < n:
        c = text[i]
        if c == '\\' and in_str:
            # エスケープシーケンス: 次の1文字ごとそのまま通す
            result.append(c)
            i += 1
            if i < n:
                result.append(text[i])
                i += 1
            continue
        if c == '"':
            in_str = not in_str
            result.append(c)
            i += 1
            continue
        if not in_str and c == '/' and i + 1 < n and text[i + 1] == '/':
            # 行末までスキップ
            while i < n and text[i] != '\n':
                i += 1
            continue
        result.append(c)
        i += 1
    return ''.join(result)


# ---------------------------------------------------------------------------
# テキストエリアのパース
# ---------------------------------------------------------------------------

def parse_pairs(text: str) -> tuple[list[tuple[str, str]], list[str]]:
    """
    "置換前=置換後" 形式のテキストをパースする。
    # 始まりはコメント行、空行はスキップ。
    戻り値: ([(old, new), ...], [エラーメッセージ, ...])
    """
    pairs  = []
    errors = []
    for i, line in enumerate(text.splitlines(), 1):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            errors.append(f"行{i}: '=' が見つかりません → {line!r}")
            continue
        old, new = line.split("=", 1)
        if not old:
            errors.append(f"行{i}: 置換前文字が空です")
            continue
        pairs.append((old, new))
    return pairs, errors


def pairs_to_text(pairs: list[dict]) -> str:
    """JSON の pairs リストをテキストエリア用の文字列に変換する。"""
    lines = []
    for p in pairs:
        old     = p.get("old", "")
        new     = p.get("new", "")
        enabled = p.get("enabled", True)
        line    = f"{old}={new}"
        if not enabled:
            line = f"# {line}"
        lines.append(line)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# メインアプリ
# ---------------------------------------------------------------------------

class ReplaceAllApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("一括文字置換ツール")
        self.root.geometry("760x700")
        self.root.resizable(True, True)
        self._build_ui()

    # ---- UI 構築 -----------------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 8, "pady": 3}

        # === フォルダ選択 ===
        frm_folder = ttk.LabelFrame(self.root, text="対象フォルダ")
        frm_folder.pack(fill=tk.X, **pad)

        self.folder_var = tk.StringVar()
        ttk.Entry(frm_folder, textvariable=self.folder_var, width=62).pack(
            side=tk.LEFT, padx=4, pady=4, fill=tk.X, expand=True)
        ttk.Button(frm_folder, text="参照...", command=self._browse_folder).pack(
            side=tk.LEFT, padx=4, pady=4)

        # === 対象オプション ===
        frm_opt = ttk.Frame(self.root)
        frm_opt.pack(fill=tk.X, **pad)

        self.recurse_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm_opt, text="サブフォルダを含める",
                        variable=self.recurse_var).pack(side=tk.LEFT)

        self.ext_vars: dict[str, tk.BooleanVar] = {}
        for ext in (".pptx", ".xlsx", ".docx"):
            v = tk.BooleanVar(value=True)
            self.ext_vars[ext] = v
            ttk.Checkbutton(frm_opt, text=ext, variable=v).pack(side=tk.LEFT, padx=6)

        # === マッチングオプション ===
        frm_match = ttk.LabelFrame(self.root, text="マッチングオプション")
        frm_match.pack(fill=tk.X, **pad)

        self.case_var     = tk.BooleanVar(value=False)   # デフォ: 区別しない
        self.fullhalf_var = tk.BooleanVar(value=False)   # デフォ: 区別しない

        ttk.Checkbutton(frm_match, text="大文字・小文字を区別する",
                        variable=self.case_var).pack(side=tk.LEFT, padx=8, pady=4)
        ttk.Checkbutton(frm_match, text="全角・半角を区別する",
                        variable=self.fullhalf_var).pack(side=tk.LEFT, padx=8, pady=4)

        # === ファイル名変更オプション ===
        frm_fname = ttk.LabelFrame(self.root, text="ファイル名変更オプション")
        frm_fname.pack(fill=tk.X, **pad)

        # 行1: 置換ペアをファイル名にも適用
        frm_fname1 = ttk.Frame(frm_fname)
        frm_fname1.pack(fill=tk.X, padx=4, pady=(4, 2))

        self.fname_replace_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_fname1, text="置換ペアをファイル名にも適用する",
                        variable=self.fname_replace_var).pack(side=tk.LEFT)

        # 行2: プレフィックス / サフィックス追加
        frm_fname2 = ttk.Frame(frm_fname)
        frm_fname2.pack(fill=tk.X, padx=4, pady=(2, 6))

        self.fname_add_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_fname2, text="ファイル名に文字列を追加する",
                        variable=self.fname_add_var,
                        command=self._on_fname_add_toggle).pack(side=tk.LEFT)

        self.fname_pos_var = tk.StringVar(value="prefix")
        self._rb_prefix = ttk.Radiobutton(frm_fname2, text="冒頭",
                                          variable=self.fname_pos_var, value="prefix")
        self._rb_suffix = ttk.Radiobutton(frm_fname2, text="末尾",
                                          variable=self.fname_pos_var, value="suffix")
        self._rb_prefix.pack(side=tk.LEFT, padx=(12, 2))
        self._rb_suffix.pack(side=tk.LEFT, padx=(0, 6))

        self.fname_str_var = tk.StringVar(value="【マスキング済】")
        self._fname_entry = ttk.Entry(frm_fname2, textvariable=self.fname_str_var, width=28)
        self._fname_entry.pack(side=tk.LEFT, padx=2)

        self._on_fname_add_toggle()   # 初期状態: 無効化

        # === 置換ペア ===
        frm_pairs = ttk.LabelFrame(
            self.root, text="置換ペア  （書式: 置換前=置換後、1行1組、# でコメント）")
        frm_pairs.pack(fill=tk.BOTH, expand=True, **pad)

        self.pairs_text = tk.Text(frm_pairs, font=("Consolas", 10), undo=True)
        sb = ttk.Scrollbar(frm_pairs, command=self.pairs_text.yview)
        self.pairs_text.config(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.pairs_text.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self.pairs_text.insert("1.0", "# 例: 旧文字=新文字\n")

        # プリセットボタン
        frm_preset = ttk.Frame(frm_pairs)
        frm_preset.pack(fill=tk.X, padx=4, pady=(0, 4))

        ttk.Button(frm_preset, text="プリセット保存",
                   command=self._save_preset, width=14).pack(side=tk.LEFT, padx=2)
        ttk.Button(frm_preset, text="プリセット読込",
                   command=self._load_preset, width=14).pack(side=tk.LEFT, padx=2)
        ttk.Button(frm_preset, text="全クリア",
                   command=self._clear_pairs, width=9).pack(side=tk.LEFT, padx=2)

        # === 実行ボタン群 ===
        frm_run = ttk.Frame(self.root)
        frm_run.pack(fill=tk.X, **pad)

        self.run_btn = ttk.Button(frm_run, text="▶ 置換実行",
                                  command=self._run_replacement)
        self.run_btn.pack(side=tk.LEFT, padx=4)
        ttk.Button(frm_run, text="ログをクリア",
                   command=self._clear_log).pack(side=tk.LEFT, padx=4)
        ttk.Button(frm_run, text="ログを保存...",
                   command=self._save_log).pack(side=tk.LEFT, padx=4)

        self.status_var = tk.StringVar(value="待機中")
        ttk.Label(frm_run, textvariable=self.status_var,
                  foreground="gray").pack(side=tk.RIGHT, padx=8)

        # === ログ ===
        frm_log = ttk.LabelFrame(self.root, text="処理ログ")
        frm_log.pack(fill=tk.BOTH, expand=True, **pad)

        self.log_text = scrolledtext.ScrolledText(
            frm_log, font=("Consolas", 9), state=tk.DISABLED, height=8)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self.log_text.tag_config("ok",    foreground="#006600")
        self.log_text.tag_config("skip",  foreground="#888800")
        self.log_text.tag_config("error", foreground="#cc0000")
        self.log_text.tag_config("info",  foreground="#000099")
        self.log_text.tag_config("head",  foreground="#000000",
                                  font=("Consolas", 9, "bold"))

    # ---- ファイル名オプション ----------------------------------------------

    def _on_fname_add_toggle(self):
        state = tk.NORMAL if self.fname_add_var.get() else tk.DISABLED
        self._rb_prefix.config(state=state)
        self._rb_suffix.config(state=state)
        self._fname_entry.config(state=state)

    # ---- ペア操作 ----------------------------------------------------------

    def _clear_pairs(self):
        self.pairs_text.delete("1.0", tk.END)

    # ---- プリセット --------------------------------------------------------

    def _save_preset(self):
        raw = self.pairs_text.get("1.0", tk.END)
        pairs_raw, errors = parse_pairs(raw)
        if errors:
            if not messagebox.askyesno("入力エラー",
                                       "\n".join(errors) + "\n\nエラー行を除いて保存しますか？"):
                return

        path = filedialog.asksaveasfilename(
            title="プリセットを保存",
            defaultextension=".json",
            filetypes=[("JSON ファイル", "*.json"), ("すべてのファイル", "*.*")],
            initialfile="preset.json",
        )
        if not path:
            return

        data = {
            "version": 1,
            "case_sensitive":     self.case_var.get(),
            "fullhalf_sensitive": self.fullhalf_var.get(),
            "pairs": [{"enabled": True, "old": o, "new": n} for o, n in pairs_raw],
        }

        # # 行（無効化ペア）も enabled:false で保存
        all_pairs = []
        for line in raw.splitlines():
            stripped = line.strip()
            if not stripped:
                continue
            if stripped.startswith("#"):
                inner = stripped[1:].strip()
                if "=" in inner:
                    o, n = inner.split("=", 1)
                    if o:
                        all_pairs.append({"enabled": False, "old": o, "new": n})
            elif "=" in stripped:
                o, n = stripped.split("=", 1)
                if o:
                    all_pairs.append({"enabled": True, "old": o, "new": n})

        data["pairs"] = all_pairs

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存完了", f"プリセットを保存しました:\n{path}")
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))

    def _load_preset(self):
        path = filedialog.askopenfilename(
            title="プリセットを読込",
            filetypes=[("JSON ファイル", "*.json"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                raw_text = f.read()
            cleaned = _strip_json_comments(raw_text)
            data = json.loads(cleaned)
        except json.JSONDecodeError as e:
            messagebox.showerror("JSON 解析エラー",
                                 f"JSON の形式が正しくありません:\n{e}")
            return
        except Exception as e:
            messagebox.showerror("読込エラー", str(e))
            return

        self.case_var.set(data.get("case_sensitive", True))
        self.fullhalf_var.set(data.get("fullhalf_sensitive", True))

        text = pairs_to_text(data.get("pairs", []))
        self.pairs_text.delete("1.0", tk.END)
        self.pairs_text.insert("1.0", text)

    # ---- ログ --------------------------------------------------------------

    def _log(self, message: str, tag: str = ""):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{ts}] {message}\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _save_log(self):
        path = filedialog.asksaveasfilename(
            title="ログを保存",
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
            initialfile=f"replace_log_{datetime.now():%Y%m%d_%H%M%S}.txt",
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.log_text.get("1.0", tk.END))
            messagebox.showinfo("保存完了", f"ログを保存しました:\n{path}")
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))

    # ---- フォルダ参照 ------------------------------------------------------

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="対象フォルダを選択")
        if folder:
            self.folder_var.set(folder)

    # ---- 置換実行 ----------------------------------------------------------

    def _run_replacement(self):
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showerror("エラー", "フォルダを選択してください")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("エラー", "有効なフォルダを指定してください")
            return

        raw = self.pairs_text.get("1.0", tk.END)
        pairs_raw, errors = parse_pairs(raw)

        if errors:
            messagebox.showerror("入力エラー", "\n".join(errors))
            return
        if not pairs_raw:
            messagebox.showerror("エラー", "置換ペアを1組以上入力してください\n"
                                           "（書式: 置換前=置換後）")
            return

        case_sensitive     = self.case_var.get()
        fullhalf_sensitive = self.fullhalf_var.get()

        replacer = build_combined_replacer(pairs_raw, case_sensitive, fullhalf_sensitive)
        if not replacer:
            messagebox.showerror("エラー", "有効な置換ペアを1組以上入力してください")
            return

        exts = [e for e, v in self.ext_vars.items() if v.get()]
        if not exts:
            messagebox.showerror("エラー", "対象拡張子を1つ以上選択してください")
            return

        self.run_btn.config(state=tk.DISABLED)
        self.status_var.set("処理中...")
        self.root.update_idletasks()

        fname_opts = {
            "apply_replace": self.fname_replace_var.get(),
            "add_string":    self.fname_add_var.get(),
            "position":      self.fname_pos_var.get(),
            "string":        self.fname_str_var.get(),
        }

        t = threading.Thread(
            target=self._do_replacement,
            args=(folder, replacer, pairs_raw, exts,
                  self.recurse_var.get(), case_sensitive, fullhalf_sensitive,
                  fname_opts),
            daemon=True,
        )
        t.start()

    def _do_replacement(self, folder, replacer, pairs_raw, exts,
                        recurse, case_sensitive, fullhalf_sensitive, fname_opts):
        start    = datetime.now()
        opt_case = "区別する" if case_sensitive else "区別しない"
        opt_fh   = "区別する" if fullhalf_sensitive else "区別しない"

        do_fname_replace = fname_opts["apply_replace"]
        do_fname_add     = fname_opts["add_string"]
        fname_position   = fname_opts["position"]    # "prefix" or "suffix"
        fname_string     = fname_opts["string"]

        self.root.after(0, lambda: self._log(f"=== 置換開始: {folder} ===", "head"))
        self.root.after(0, lambda: self._log(
            f"大文字小文字: {opt_case} / 全角半角: {opt_fh}", "info"))
        self.root.after(0, lambda: self._log(
            f"置換ペア ({len(pairs_raw)}組): " +
            " / ".join(f'"{o}"→"{n}"' for o, n in pairs_raw), "info"))
        if do_fname_replace:
            self.root.after(0, lambda: self._log("ファイル名: 置換ペアを適用", "info"))
        if do_fname_add and fname_string:
            pos_label = "冒頭" if fname_position == "prefix" else "末尾"
            self.root.after(0, lambda: self._log(
                f"ファイル名: {pos_label}に「{fname_string}」を追加", "info"))

        total_files        = 0
        changed_files      = 0
        total_replacements = 0
        renamed_files      = 0
        error_count        = 0

        walker = os.walk(folder) if recurse else [(folder, [], os.listdir(folder))]

        for dirpath, _, filenames in walker:
            for fname in sorted(filenames):
                stem, ext = os.path.splitext(fname)
                if ext.lower() not in exts:
                    continue
                total_files += 1
                fpath = os.path.join(dirpath, fname)
                rel   = os.path.relpath(fpath, folder)

                # --- 内容の置換 ---
                handler = HANDLERS[ext.lower()]
                content_n = 0
                try:
                    content_n = handler(fpath, replacer)
                    if content_n:
                        changed_files      += 1
                        total_replacements += content_n
                        msg, tag = f"[変更]    {rel}  ({content_n}箇所)", "ok"
                    else:
                        msg, tag = f"[変更なし] {rel}", "skip"
                except Exception as e:
                    error_count += 1
                    msg, tag = f"[エラー]  {rel}: {e}", "error"

                self.root.after(0, lambda m=msg, t=tag: self._log(m, t))

                # --- ファイル名の変更 ---
                new_stem = stem
                if do_fname_replace:
                    new_stem = replacer(new_stem)
                if do_fname_add and fname_string:
                    if fname_position == "prefix":
                        if not new_stem.startswith(fname_string):
                            new_stem = fname_string + new_stem
                    else:
                        if not new_stem.endswith(fname_string):
                            new_stem = new_stem + fname_string

                if new_stem != stem:
                    new_fname = new_stem + ext
                    new_fpath = os.path.join(dirpath, new_fname)
                    try:
                        if os.path.exists(new_fpath):
                            rmsg = f"[名前変更スキップ] {rel} → {new_fname}（同名ファイルが存在）"
                            self.root.after(0, lambda m=rmsg: self._log(m, "error"))
                        else:
                            os.rename(fpath, new_fpath)
                            renamed_files += 1
                            new_rel = os.path.relpath(new_fpath, folder)
                            rmsg = f"[名前変更]  {rel} → {new_rel}"
                            self.root.after(0, lambda m=rmsg: self._log(m, "ok"))
                    except Exception as e:
                        error_count += 1
                        rmsg = f"[名前変更エラー] {rel}: {e}"
                        self.root.after(0, lambda m=rmsg: self._log(m, "error"))

        elapsed = (datetime.now() - start).total_seconds()
        summary = (
            f"=== 完了 ({elapsed:.1f}秒)  "
            f"対象:{total_files}件  内容変更:{changed_files}件  "
            f"置換:{total_replacements}箇所  名前変更:{renamed_files}件  "
            f"エラー:{error_count}件 ==="
        )
        self.root.after(0, lambda: self._log(summary, "head"))
        self.root.after(0, lambda: self.status_var.set(
            f"完了 — 内容:{changed_files}/{total_files}件変更  名前:{renamed_files}件変更"))
        self.root.after(0, lambda: self.run_btn.config(state=tk.NORMAL))


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------

def _check_deps():
    missing = []
    for pkg, mod in [("python-pptx", "pptx"),
                     ("openpyxl",    "openpyxl"),
                     ("python-docx", "docx")]:
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    return missing


if __name__ == "__main__":
    missing = _check_deps()
    if missing:
        print("以下のライブラリが不足しています。インストールしてください:")
        print(f"  pip install {' '.join(missing)}")
        sys.exit(1)

    root = tk.Tk()
    ReplaceAllApp(root)
    root.mainloop()
