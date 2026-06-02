import sqlite3
import random
import os
import pandas as pd
from dataclasses import dataclass
from copy import deepcopy
from typing import cast
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# -----------------------------
# Config
# -----------------------------

ITERATIONS = 7000
TEAM_A_NAME = "Light Team"
TEAM_B_NAME = "Dark Team"
TEAM_A_JERSEY = "LIGHT"
TEAM_B_JERSEY = "DARK"
OUTPUT_FILE = "game_night_teams.xlsx"
DB_FILE = "hockey_players.db"

POSITIONS = ["F", "D", "F/D"]

# -----------------------------
# Database
# -----------------------------

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS players (
                id      INTEGER PRIMARY KEY AUTOINCREMENT,
                name    TEXT    NOT NULL UNIQUE,
                rank    INTEGER NOT NULL,
                position TEXT   NOT NULL
            )
        """)
        conn.commit()


def db_get_all_players():
    with get_db() as conn:
        return conn.execute("SELECT * FROM players ORDER BY name").fetchall()


def db_add_player(name, rank, position):
    with get_db() as conn:
        conn.execute(
            "INSERT INTO players (name, rank, position) VALUES (?, ?, ?)",
            (name.strip(), int(rank), position)
        )
        conn.commit()


def db_update_player(player_id, name, rank, position):
    with get_db() as conn:
        conn.execute(
            "UPDATE players SET name=?, rank=?, position=? WHERE id=?",
            (name.strip(), int(rank), position, player_id)
        )
        conn.commit()


def db_delete_player(player_id):
    with get_db() as conn:
        conn.execute("DELETE FROM players WHERE id=?", (player_id,))
        conn.commit()


def db_import_from_excel(filepath):
    """
    Import players from an Excel file into the database.
    Expects columns: Name, Rank, Position.
    Skips rows where the name already exists (no overwrite).
    Returns (imported_count, skipped_count, errors).
    """
    df = pd.read_excel(filepath)
    required = {"Name", "Rank", "Position"}
    if not required.issubset(df.columns):
        raise ValueError(f"Excel file must contain columns: {', '.join(required)}")

    imported, skipped, errors = 0, 0, []

    with get_db() as conn:
        for _, row in df.iterrows():
            name     = str(row["Name"]).strip()
            position = str(row["Position"]).strip().upper()
            try:
                rank = int(row["Rank"])
                if not name or name.lower() == "nan":
                    continue
                if position not in ("F", "D", "F/D"):
                    errors.append(f"{name}: invalid position '{position}'")
                    continue
                conn.execute(
                    "INSERT INTO players (name, rank, position) VALUES (?, ?, ?)",
                    (name, rank, position)
                )
                imported += 1
            except sqlite3.IntegrityError:
                skipped += 1          # already exists — leave it alone
            except Exception as e:
                errors.append(f"{name}: {e}")
        conn.commit()

    return imported, skipped, errors


# -----------------------------
# Player dataclass
# -----------------------------

@dataclass
class Player:
    name: str
    rank: int
    position: str


# -----------------------------
# Balancer logic
# -----------------------------

def empty_team():
    return {"F": [], "D": [], "total": 0}


def can_add(team, pos, limits):
    return len(team[pos]) < limits.get(pos, 0)


def compute_limits(players):
    f_count  = sum(1 for p in players if p.position == "F")
    d_count  = sum(1 for p in players if p.position == "D")
    flex     = sum(1 for p in players if p.position == "F/D")
    flex_f   = flex // 2
    flex_d   = flex - flex_f
    total_f  = f_count + flex_f
    total_d  = d_count + flex_d
    return (
        {"F": (total_f + 1) // 2, "D": (total_d + 1) // 2},
        {"F": total_f // 2,       "D": total_d // 2}
    )


def assign_player(team, player, pos):
    team[pos].append(player)
    team["total"] += player.rank


def attempt_build(players_list, limits_a, limits_b):
    random.shuffle(players_list)
    team_a, team_b = empty_team(), empty_team()

    for player in players_list:
        options = []
        for team, limits in [(team_a, limits_a), (team_b, limits_b)]:
            if player.position == "F":
                if can_add(team, "F", limits): options.append((team, "F"))
            elif player.position == "D":
                if can_add(team, "D", limits): options.append((team, "D"))
            else:
                if can_add(team, "F", limits): options.append((team, "F"))
                if can_add(team, "D", limits): options.append((team, "D"))

        if not options:
            for team, limits in [(team_a, limits_a), (team_b, limits_b)]:
                for pos in ["F", "D"]:
                    if can_add(team, pos, limits): options.append((team, pos))

        if not options:
            options = [(team_a, "F"), (team_b, "F"), (team_a, "D"), (team_b, "D")]

        best_move, best_diff = (None, None), float("inf")
        for team, pos in options:
            pa = team_a["total"] + (player.rank if team is team_a else 0)
            pb = team_b["total"] + (player.rank if team is team_b else 0)
            if abs(pa - pb) < best_diff:
                best_diff = abs(pa - pb)
                best_move = (team, pos)

        assign_player(best_move[0], player, best_move[1])

    return team_a, team_b


def build_teams(players):
    if len(players) < 10:
        raise ValueError("Not enough selected players (minimum 10)")

    limits_a, limits_b = compute_limits(players)
    best, best_diff = (None, None), float("inf")

    for _ in range(ITERATIONS):
        a, b = attempt_build(players.copy(), limits_a, limits_b)
        diff = abs(a["total"] - b["total"])
        if diff < best_diff:
            best_diff = diff
            best = (deepcopy(a), deepcopy(b))
        if best_diff == 0:
            break

    return best[0], best[1], best_diff


def export_workbook(team_a, team_b, diff):
    def build_df(team, jersey):
        return pd.DataFrame([
            {"Name": p.name, "Rank": p.rank, "Position": pos, "Jersey": jersey}
            for pos in ["F", "D"] for p in team[pos]
        ])

    summary = pd.DataFrame({
        "Metric": [
            "Light Team Forwards", "Light Team Defence", "Light Team Total Players", "Light Team Total Rank",
            "Dark Team Forwards",  "Dark Team Defence",  "Dark Team Total Players",  "Dark Team Total Rank",
            "Skill Difference"
        ],
        "Value": [
            len(team_a["F"]), len(team_a["D"]), len(team_a["F"]) + len(team_a["D"]), team_a["total"],
            len(team_b["F"]), len(team_b["D"]), len(team_b["F"]) + len(team_b["D"]), team_b["total"],
            diff
        ]
    })

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        build_df(team_a, TEAM_A_JERSEY).to_excel(writer, sheet_name=TEAM_A_NAME, index=False)
        build_df(team_b, TEAM_B_JERSEY).to_excel(writer, sheet_name=TEAM_B_NAME, index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)


# ==============================
# GUI
# ==============================

class HockeyApp(tk.Tk):

    # colour palette
    BG       = "#0d1117"
    SURFACE  = "#161b22"
    BORDER   = "#30363d"
    ICE      = "#cae8ff"
    ACCENT   = "#388bfd"
    RED      = "#f85149"
    GREEN    = "#3fb950"
    TEXT     = "#e6edf3"
    SUBTEXT  = "#8b949e"
    FONT     = "Helvetica"

    def __init__(self):
        super().__init__()
        self.title("🏒 Hockey Team Balancer")
        self.geometry("860x600")
        self.minsize(760, 520)
        self.configure(bg=self.BG)

        self._apply_styles()
        self._build_ui()
        init_db()
        self.refresh_roster()

    # ------------------------------------------------------------------
    # Styles
    # ------------------------------------------------------------------

    def _apply_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")

        s.configure(".", background=self.BG, foreground=self.TEXT,
                    font=(self.FONT, 11), borderwidth=0)

        s.configure("TNotebook", background=self.BG, borderwidth=0, tabmargins=0)
        s.configure("TNotebook.Tab",
                    background=self.SURFACE, foreground=self.SUBTEXT,
                    padding=[18, 8], font=(self.FONT, 11, "bold"))
        s.map("TNotebook.Tab",
              background=[("selected", self.BG)],
              foreground=[("selected", self.ICE)])

        s.configure("Treeview",
                    background=self.SURFACE, foreground=self.TEXT,
                    fieldbackground=self.SURFACE, rowheight=28,
                    font=(self.FONT, 11))
        s.configure("Treeview.Heading",
                    background=self.BORDER, foreground=self.ICE,
                    font=(self.FONT, 11, "bold"))
        s.map("Treeview", background=[("selected", self.ACCENT)])

        s.configure("TScrollbar", background=self.BORDER,
                    troughcolor=self.SURFACE, arrowcolor=self.SUBTEXT)

        s.configure("TCombobox", fieldbackground=self.SURFACE,
                    background=self.SURFACE, foreground=self.TEXT,
                    arrowcolor=self.ICE)
        s.map("TCombobox", fieldbackground=[("readonly", self.SURFACE)])

        self.option_add("*TCombobox*Listbox.background", self.SURFACE)
        self.option_add("*TCombobox*Listbox.foreground", self.TEXT)
        self.option_add("*TCombobox*Listbox.selectBackground", self.ACCENT)

    def _btn(self, parent, text, cmd, color=None, **kw):
        bg = color or self.ACCENT
        b = tk.Button(parent, text=text, command=cmd,
                      bg=bg, fg=self.BG if bg != self.RED else "white",
                      activebackground=bg, activeforeground=self.BG,
                      relief="flat", cursor="hand2",
                      font=(self.FONT, 11, "bold"),
                      padx=14, pady=6, **kw)
        return b

    def _entry(self, parent, width=22):
        e = tk.Entry(parent, width=width,
                     bg=self.SURFACE, fg=self.TEXT,
                     insertbackground=self.TEXT,
                     relief="flat", font=(self.FONT, 11),
                     highlightthickness=1,
                     highlightbackground=self.BORDER,
                     highlightcolor=self.ACCENT)
        return e

    def _label(self, parent, text, size=11, color=None, bold=False):
        font = (self.FONT, size, "bold") if bold else (self.FONT, size)
        return tk.Label(parent, text=text, bg=self.BG,
                        fg=color or self.TEXT, font=font)

    # ------------------------------------------------------------------
    # Main layout
    # ------------------------------------------------------------------

    def _build_ui(self):
        header = tk.Frame(self, bg=self.BG)
        header.pack(fill="x", padx=24, pady=(18, 4))
        tk.Label(header, text="🏒  Hockey Team Balancer",
                 bg=self.BG, fg=self.ICE,
                 font=(self.FONT, 20, "bold")).pack(side="left")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=16, pady=(8, 16))

        self.tab_players  = tk.Frame(nb, bg=self.BG)
        self.tab_gameday  = tk.Frame(nb, bg=self.BG)
        self.tab_balance  = tk.Frame(nb, bg=self.BG)

        nb.add(self.tab_players, text="  👥  Players  ")
        nb.add(self.tab_gameday, text="  🏒  Game Day  ")
        nb.add(self.tab_balance, text="  ⚖️   Balance Teams  ")

        nb.bind("<<NotebookTabChanged>>", self._on_tab_change)

        self._build_players_tab()
        self._build_gameday_tab()
        self._build_balance_tab()

    def _on_tab_change(self, event):
        tab = event.widget.index("current")
        if tab == 1:
            self.refresh_gameday()
        elif tab == 2:
            self.refresh_balance_list()

    # ------------------------------------------------------------------
    # Tab 1 — Players
    # ------------------------------------------------------------------

    def _build_players_tab(self):
        p = self.tab_players

        # ---- roster tree ----
        tree_frame = tk.Frame(p, bg=self.BG)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(16, 8))

        cols = ("Name", "Rank", "Position")
        self.roster_tree = ttk.Treeview(tree_frame, columns=cols,
                                        show="headings", selectmode="browse")
        for col, w in zip(cols, (340, 80, 100)):
            self.roster_tree.heading(col, text=col)
            self.roster_tree.column(col, width=w, anchor="w" if col == "Name" else "center")

        sb = ttk.Scrollbar(tree_frame, orient="vertical",
                           command=self.roster_tree.yview)
        self.roster_tree.configure(yscrollcommand=sb.set)
        self.roster_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.roster_tree.bind("<<TreeviewSelect>>", self._on_roster_select)

        # ---- form ----
        form = tk.Frame(p, bg=self.BG)
        form.pack(fill="x", padx=20, pady=(0, 16))

        self._label(form, "Name").grid(row=0, column=0, sticky="w", padx=(0,6))
        self.entry_name = self._entry(form, 28)
        self.entry_name.grid(row=0, column=1, padx=(0, 18))

        self._label(form, "Rank (1–99)").grid(row=0, column=2, sticky="w", padx=(0,6))
        self.entry_rank = self._entry(form, 6)
        self.entry_rank.grid(row=0, column=3, padx=(0, 18))

        self._label(form, "Position").grid(row=0, column=4, sticky="w", padx=(0,6))
        self.pos_var = tk.StringVar(value="F")
        self.pos_combo = ttk.Combobox(form, textvariable=self.pos_var,
                                      values=POSITIONS, state="readonly", width=7)
        self.pos_combo.grid(row=0, column=5, padx=(0, 18))

        btn_row = tk.Frame(p, bg=self.BG)
        btn_row.pack(fill="x", padx=20, pady=(0, 12))

        self._btn(btn_row, "+ Add Player",    self._add_player,    self.GREEN).pack(side="left", padx=(0,8))
        self._btn(btn_row, "✎ Save Changes",  self._update_player, self.ACCENT).pack(side="left", padx=(0,8))
        self._btn(btn_row, "✕ Delete Player", self._delete_player, self.RED  ).pack(side="left", padx=(0,8))
        self._btn(btn_row, "📥 Import from Excel", self._import_excel, self.SURFACE).pack(side="right")

        self._selected_player_id = None

    def refresh_roster(self):
        self.roster_tree.delete(*self.roster_tree.get_children())
        for row in db_get_all_players():
            self.roster_tree.insert("", "end", iid=str(row["id"]),
                                    values=(row["name"], row["rank"], row["position"]))

    def _on_roster_select(self, _event):
        sel = self.roster_tree.selection()
        if not sel:
            return
        vals = self.roster_tree.item(sel[0], "values")
        self._selected_player_id = int(sel[0])
        self.entry_name.delete(0, "end");  self.entry_name.insert(0, vals[0])
        self.entry_rank.delete(0, "end");  self.entry_rank.insert(0, vals[1])
        self.pos_var.set(vals[2])

    def _validate_form(self):
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showerror("Validation", "Name cannot be empty."); return None, None, None
        try:
            rank = int(self.entry_rank.get())
            assert 1 <= rank <= 99
        except:
            messagebox.showerror("Validation", "Rank must be a number between 1 and 99."); return None, None, None
        return name, rank, self.pos_var.get()

    def _add_player(self):
        name, rank, pos = self._validate_form()
        if name is None: return
        try:
            db_add_player(name, rank, pos)
            self.refresh_roster()
            self.entry_name.delete(0, "end")
            self.entry_rank.delete(0, "end")
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"A player named '{name}' already exists.")

    def _update_player(self):
        if not self._selected_player_id:
            messagebox.showwarning("No Selection", "Select a player to edit."); return
        name, rank, pos = self._validate_form()
        if name is None: return
        db_update_player(self._selected_player_id, name, rank, pos)
        self.refresh_roster()

    def _delete_player(self):
        if not self._selected_player_id:
            messagebox.showwarning("No Selection", "Select a player to delete."); return
        vals = self.roster_tree.item(str(self._selected_player_id), "values")
        if not messagebox.askyesno("Confirm", f"Delete {vals[0]}?"):
            return
        db_delete_player(self._selected_player_id)
        self._selected_player_id = None
        self.entry_name.delete(0, "end")
        self.entry_rank.delete(0, "end")
        self.refresh_roster()

    def _import_excel(self):
        filepath = filedialog.askopenfilename(
            title="Select players Excel file",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not filepath:
            return
        try:
            imported, skipped, errors = db_import_from_excel(filepath)
        except ValueError as e:
            messagebox.showerror("Import Error", str(e))
            return
        except Exception as e:
            messagebox.showerror("Import Error", f"Could not read file:\n{e}")
            return

        self.refresh_roster()

        msg = f"✔ Import complete\n\n{imported} players imported"
        if skipped:
            msg += f"\n{skipped} skipped (already in database)"
        if errors:
            msg += f"\n\n⚠ {len(errors)} row(s) had errors:\n" + "\n".join(errors[:10])
            if len(errors) > 10:
                msg += f"\n…and {len(errors)-10} more"

        messagebox.showinfo("Import from Excel", msg)

    # ------------------------------------------------------------------
    # Tab 2 — Game Day
    # ------------------------------------------------------------------

    def _build_gameday_tab(self):
        g = self.tab_gameday

        top = tk.Frame(g, bg=self.BG)
        top.pack(fill="x", padx=20, pady=(14, 6))
        self._label(top, "Select players available for tonight's game",
                    size=12, bold=True).pack(side="left")

        btn_row = tk.Frame(g, bg=self.BG)
        btn_row.pack(fill="x", padx=20, pady=(0, 8))
        self._btn(btn_row, "✔ Select All",   self._gameday_select_all,   self.ACCENT).pack(side="left", padx=(0,8))
        self._btn(btn_row, "✕ Deselect All", self._gameday_deselect_all, self.SURFACE).pack(side="left")

        tree_frame = tk.Frame(g, bg=self.BG)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        cols = ("Playing", "Name", "Rank", "Position")
        self.gameday_tree = ttk.Treeview(tree_frame, columns=cols,
                                         show="headings", selectmode="browse")
        widths = {"Playing": 70, "Name": 300, "Rank": 80, "Position": 100}
        for col in cols:
            self.gameday_tree.heading(col, text=col)
            self.gameday_tree.column(col, width=widths[col],
                                     anchor="center" if col != "Name" else "w")

        sb = ttk.Scrollbar(tree_frame, orient="vertical",
                           command=self.gameday_tree.yview)
        self.gameday_tree.configure(yscrollcommand=sb.set)
        self.gameday_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.gameday_tree.tag_configure("playing",     background="#1a2f1a", foreground=self.GREEN)
        self.gameday_tree.tag_configure("not_playing", background=self.SURFACE, foreground=self.SUBTEXT)
        self.gameday_tree.bind("<ButtonRelease-1>", self._toggle_playing)
        self.gameday_tree.bind("<space>",           self._toggle_playing)

        # track which player ids are "in"
        self._playing = set()

        status_bar = tk.Frame(g, bg=self.SURFACE)
        status_bar.pack(fill="x", padx=20, pady=(0, 4))
        self.gameday_count_lbl = tk.Label(status_bar, text="0 players selected",
                                          bg=self.SURFACE, fg=self.ICE,
                                          font=(self.FONT, 11), padx=10, pady=4)
        self.gameday_count_lbl.pack(side="left")

    def refresh_gameday(self):
        self.gameday_tree.delete(*self.gameday_tree.get_children())
        for row in db_get_all_players():
            pid   = str(row["id"])
            check = "✔" if int(row["id"]) in self._playing else "–"
            tag   = "playing" if int(row["id"]) in self._playing else "not_playing"
            self.gameday_tree.insert("", "end", iid=pid,
                                     values=(check, row["name"], row["rank"], row["position"]),
                                     tags=(tag,))
        self._update_gameday_count()

    def _toggle_playing(self, _event):
        sel = self.gameday_tree.selection()
        if not sel: return
        pid = int(sel[0])
        if pid in self._playing:
            self._playing.discard(pid)
        else:
            self._playing.add(pid)
        self.refresh_gameday()
        # keep selection
        self.gameday_tree.selection_set(str(pid))

    def _gameday_select_all(self):
        self._playing = {int(row["id"]) for row in db_get_all_players()}
        self.refresh_gameday()

    def _gameday_deselect_all(self):
        self._playing.clear()
        self.refresh_gameday()

    def _update_gameday_count(self):
        n = len(self._playing)
        color = self.GREEN if n >= 10 else self.RED
        self.gameday_count_lbl.config(
            text=f"{n} player{'s' if n != 1 else ''} selected for tonight",
            fg=color
        )

    # ------------------------------------------------------------------
    # Tab 3 — Balance Teams
    # ------------------------------------------------------------------

    def _build_balance_tab(self):
        b = self.tab_balance

        top = tk.Frame(b, bg=self.BG)
        top.pack(fill="x", padx=20, pady=(14, 6))
        self._label(top, "Balance Teams", size=12, bold=True).pack(side="left")

        # summary of who is playing
        self.balance_players_frame = tk.Frame(b, bg=self.SURFACE,
                                              highlightthickness=1,
                                              highlightbackground=self.BORDER)
        self.balance_players_frame.pack(fill="x", padx=20, pady=(0, 12))
        self.balance_summary_lbl = tk.Label(self.balance_players_frame,
                                            text="Go to Game Day tab to select players.",
                                            bg=self.SURFACE, fg=self.SUBTEXT,
                                            font=(self.FONT, 11), padx=12, pady=8,
                                            anchor="w", justify="left")
        self.balance_summary_lbl.pack(fill="x")

        self._btn(b, "⚖️  Generate Balanced Teams", self._run_balance,
                  self.ACCENT).pack(pady=(0, 16))

        self.balance_status = tk.Label(b, text="", bg=self.BG,
                                       fg=self.GREEN, font=(self.FONT, 12, "bold"))
        self.balance_status.pack()

        # results side by side
        results = tk.Frame(b, bg=self.BG)
        results.pack(fill="both", expand=True, padx=20, pady=(8, 16))
        results.columnconfigure(0, weight=1)
        results.columnconfigure(1, weight=1)

        for col, label, attr in [(0, "🤍 Light Team", "light_tree"),
                                  (1, "🖤 Dark Team",  "dark_tree")]:
            frame = tk.Frame(results, bg=self.BG)
            frame.grid(row=0, column=col, sticky="nsew", padx=(0, 8) if col == 0 else (8, 0))
            tk.Label(frame, text=label, bg=self.BG, fg=self.ICE,
                     font=(self.FONT, 12, "bold")).pack(anchor="w", pady=(0, 4))

            cols = ("Name", "Rank", "Pos")
            tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
            for c, w in zip(cols, (180, 60, 60)):
                tree.heading(c, text=c)
                tree.column(c, width=w, anchor="w" if c == "Name" else "center")
            tree.pack(fill="both", expand=True)
            setattr(self, attr, tree)

    def refresh_balance_list(self):
        rows = db_get_all_players()
        playing = [r for r in rows if int(r["id"]) in self._playing]
        n = len(playing)
        f = sum(1 for r in playing if r["position"] == "F")
        d = sum(1 for r in playing if r["position"] == "D")
        x = sum(1 for r in playing if r["position"] == "F/D")
        color = self.GREEN if n >= 10 else self.RED
        self.balance_summary_lbl.config(
            text=f"{n} players selected  •  {f} Forwards  •  {d} Defence  •  {x} Flex",
            fg=color
        )

    def _run_balance(self):
        rows = db_get_all_players()
        players = [
            Player(name=r["name"], rank=r["rank"], position=r["position"])
            for r in rows if int(r["id"]) in self._playing
        ]
        try:
            team_a, team_b, diff = build_teams(players)
            export_workbook(team_a, team_b, diff)
        except ValueError as e:
            messagebox.showerror("Error", str(e)); return
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error:\n{e}"); return

        self.balance_status.config(
            text=f"✔  Teams balanced!  Skill difference: {diff}   →   Saved to {OUTPUT_FILE}",
            fg=self.GREEN
        )

        for tree, team in [(self.light_tree, team_a), (self.dark_tree, team_b)]:
            tree.delete(*tree.get_children())
            for pos in ["F", "D"]:
                for p in team[pos]:
                    tree.insert("", "end", values=(p.name, p.rank, pos))


# ==============================
# Entry point
# ==============================

if __name__ == "__main__":
    app = HockeyApp()
    app.mainloop()
