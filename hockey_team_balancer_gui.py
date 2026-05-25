import pandas as pd
import random
from dataclasses import dataclass
from copy import deepcopy
import tkinter as tk
from typing import cast
from tkinter import filedialog, messagebox


# -----------------------------
# Player Class
# -----------------------------

@dataclass
class Player:
    name: str
    rank: int
    position: str


# -----------------------------
# Config
# -----------------------------

ITERATIONS = 7000

TEAM_A_NAME = "Light Team"
TEAM_B_NAME = "Dark Team"

TEAM_A_JERSEY = "LIGHT"
TEAM_B_JERSEY = "DARK"

OUTPUT_FILE = "game_night_teams.xlsx"


# -----------------------------
# Helpers
# -----------------------------

def is_selected(value):
    if pd.isna(value):
        return False
    return str(value).strip().upper() in ["TRUE", "1", "YES", "Y"]


def empty_team():
    return {"F": [], "D": [], "total": 0}


def can_add(team, pos, limits):
    if pos == "F":
        return len(team["F"]) < limits["F"]
    if pos == "D":
        return len(team["D"]) < limits["D"]
    return False


def compute_limits(players):
    """
    Dynamically compute per-team slot limits based on actual player counts.
    For each position group, split as evenly as possible.
    If odd, team_a gets the extra player.
    F/D flex players are distributed across both position slots.
    """
    f_count = sum(1 for p in players if p.position == "F")
    d_count = sum(1 for p in players if p.position == "D")
    flex_count = sum(1 for p in players if p.position == "F/D")

    # Distribute flex players evenly across F and D slots
    flex_to_f = flex_count // 2
    flex_to_d = flex_count - flex_to_f

    total_f_slots = f_count + flex_to_f
    total_d_slots = d_count + flex_to_d

    # Split slots between the two teams; team_a gets the extra if odd
    a_f = (total_f_slots + 1) // 2
    b_f = total_f_slots // 2
    a_d = (total_d_slots + 1) // 2
    b_d = total_d_slots // 2

    limits_a = {"F": a_f, "D": a_d}
    limits_b = {"F": b_f, "D": b_d}

    return limits_a, limits_b


def assign_player(team, player, pos):
    team[pos].append(player)
    team["total"] += player.rank


# -----------------------------
# Core Optimizer
# -----------------------------

def attempt_build(players_list, limits_a, limits_b):

    random.shuffle(players_list)

    team_a = empty_team()
    team_b = empty_team()

    for player in players_list:

        options = []

        for team, limits in [(team_a, limits_a), (team_b, limits_b)]:

            if player.position == "F":
                if can_add(team, "F", limits):
                    options.append((team, "F"))

            elif player.position == "D":
                if can_add(team, "D", limits):
                    options.append((team, "D"))

            else:  # FLEX
                if can_add(team, "F", limits):
                    options.append((team, "F"))
                if can_add(team, "D", limits):
                    options.append((team, "D"))

        if not options:
            # Overflow safety: place anywhere with remaining capacity
            for team, limits in [(team_a, limits_a), (team_b, limits_b)]:
                for pos in ["F", "D"]:
                    if can_add(team, pos, limits):
                        options.append((team, pos))

        if not options:
            # All slots full (shouldn't happen with correct limits, but be safe)
            options = [(team_a, "F"), (team_b, "F"), (team_a, "D"), (team_b, "D")]

        best_move = (None, None)
        best_diff = float("inf")

        for team, pos in options:

            proj_a = team_a["total"]
            proj_b = team_b["total"]

            if team == team_a:
                proj_a += player.rank
            else:
                proj_b += player.rank

            diff = abs(proj_a - proj_b)

            if diff < best_diff:
                best_diff = diff
                best_move = (team, pos)

        assign_player(best_move[0], player, best_move[1])

    return team_a, team_b


# -----------------------------
# Run Full Build
# -----------------------------

def build_teams(input_file):

    df = pd.read_excel(input_file)

    required_columns = {"Selected", "Name", "Rank", "Position"}

    if not required_columns.issubset(df.columns):
        raise ValueError("Excel must contain Selected, Name, Rank, Position")

    players = []

    for _, row in df.iterrows():

        if not is_selected(row["Selected"]):
            continue

        players.append(
            Player(
                name=cast(str, row["Name"]),
                rank=int(cast(int, row["Rank"])),
                position=cast(str, row["Position"]).strip().upper()
            )
        )

    if len(players) < 10:
        raise ValueError("Not enough selected players")

    limits_a, limits_b = compute_limits(players)

    best = (None, None)
    best_diff = float("inf")

    for _ in range(ITERATIONS):

        a, b = attempt_build(players.copy(), limits_a, limits_b)
        diff = abs(a["total"] - b["total"])

        if diff < best_diff:
            best_diff = diff
            best = (deepcopy(a), deepcopy(b))

        if best_diff == 0:
            break

    return best[0], best[1], best_diff, limits_a, limits_b


# -----------------------------
# Export Workbook
# -----------------------------

def export_workbook(team_a, team_b, diff, limits_a, limits_b):

    def build_df(team, jersey):

        rows = []

        for pos in ["F", "D"]:
            for p in team[pos]:
                rows.append({
                    "Name": p.name,
                    "Rank": p.rank,
                    "Position": pos,
                    "Jersey": jersey
                })

        return pd.DataFrame(rows)

    df_light = build_df(team_a, TEAM_A_JERSEY)
    df_dark = build_df(team_b, TEAM_B_JERSEY)

    summary = {
        "Metric": [
            "Light Team Forwards",
            "Light Team Defence",
            "Light Team Total Players",
            "Light Team Total Rank",
            "Dark Team Forwards",
            "Dark Team Defence",
            "Dark Team Total Players",
            "Dark Team Total Rank",
            "Skill Difference"
        ],
        "Value": [
            len(team_a["F"]),
            len(team_a["D"]),
            len(team_a["F"]) + len(team_a["D"]),
            team_a["total"],
            len(team_b["F"]),
            len(team_b["D"]),
            len(team_b["F"]) + len(team_b["D"]),
            team_b["total"],
            diff
        ]
    }

    df_summary = pd.DataFrame(summary)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

        df_light.to_excel(writer, sheet_name=TEAM_A_NAME, index=False)
        df_dark.to_excel(writer, sheet_name=TEAM_B_NAME, index=False)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)


# -----------------------------
# GUI
# -----------------------------

def select_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if file_path:
        file_label.config(text=file_path)
        generate_button.config(state="normal")


def generate_teams():

    try:
        input_file = file_label.cget("text")

        team_a, team_b, diff, limits_a, limits_b = build_teams(input_file)
        export_workbook(team_a, team_b, diff, limits_a, limits_b)

        status_label.config(
            text=f"✔ Teams Created Successfully!\nSkill Difference: {diff}",
            fg="green"
        )

        messagebox.showinfo(
            "Success",
            f"Workbook created:\n{OUTPUT_FILE}"
        )

    except Exception as e:

        status_label.config(text="❌ Error Occurred", fg="red")
        messagebox.showerror("Error", str(e))


# -----------------------------
# Build Window
# -----------------------------

root = tk.Tk()
root.title("Hockey Team Balancer")
root.geometry("460x300")
root.resizable(False, False)

title = tk.Label(root, text="🏒 Hockey Team Balancer", font=("Arial", 16, "bold"))
title.pack(pady=10)

select_button = tk.Button(root, text="Select Excel File", command=select_file, width=25)
select_button.pack(pady=5)

file_label = tk.Label(root, text="No file selected", wraplength=420)
file_label.pack(pady=5)

generate_button = tk.Button(
    root,
    text="Generate Balanced Teams",
    command=generate_teams,
    width=25,
    state="disabled"
)
generate_button.pack(pady=15)

status_label = tk.Label(root, text="", font=("Arial", 11))
status_label.pack(pady=10)

root.mainloop()