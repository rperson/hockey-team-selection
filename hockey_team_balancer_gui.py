import pandas as pd
import random
from dataclasses import dataclass
from copy import deepcopy
import tkinter as tk
from typing import cast
from tkinter import filedialog, messagebox
import email
from email.message import EmailMessage # Added
from email import policy
import re


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

FORWARDS_TARGET = 6
DEFENCE_TARGET = 4
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


def can_add(team, pos):
    if pos == "F":
        return len(team["F"]) < FORWARDS_TARGET
    if pos == "D":
        return len(team["D"]) < DEFENCE_TARGET
    return False

def extract_players_from_eml(eml_file_path):
    """
    Parses an EML file to extract the list of player names from the final roster section.
    """
    with open(eml_file_path, 'rb') as fp:
        msg: EmailMessage = email.message_from_binary_file(fp, policy=policy.default) # Added type hint

    player_names = []
    payload = ""
    
    # Attempt to get the plain text part of the email
    if msg.is_multipart():
        for part in msg.iter_parts():
            if part.get_content_type() == 'text/plain':
                payload = part.get_payload(decode=True) # Removed redundant .decode()
                break
        else:
            raise ValueError("No plain text part found in the EML file.")
    else:
        if msg.get_content_type() == 'text/plain':
            payload = msg.get_payload(decode=True) # Removed redundant .decode()
        else:
            raise ValueError("EML file is not plain text or multipart with a plain text part.")

    # Find the start of the roster list in the text
    roster_start_marker = "Which leaves you with a roster of"
    lines = payload.splitlines()
    
    in_roster_section = False
    for line in lines:
        stripped_line = line.strip()
        if stripped_line.startswith(roster_start_marker):
            in_roster_section = True
            continue # Skip the marker line itself
        
        if in_roster_section:
            if stripped_line: # If the line is not empty
                # Clean up potential extra spaces or non-breaking spaces
                clean_name = re.sub(r'[\s\xa0]+', ' ', stripped_line).strip()
                if clean_name: # Ensure the name is not just an empty string after cleaning
                    player_names.append(clean_name)
            else: # If an empty line is encountered, we've passed the roster list
                break
                
    if not player_names:
        raise ValueError("Could not find the player roster in the EML file or it was empty.")

    return player_names

def assign_player(team, player, pos):
    team[pos].append(player)
    team["total"] += player.rank


# -----------------------------
# Core Optimizer
# -----------------------------

def attempt_build(players_list):

    random.shuffle(players_list)

    team_a = empty_team()
    team_b = empty_team()

    for player in players_list:

        options = []

        for team in [team_a, team_b]:

            if player.position == "F":
                if can_add(team, "F"):
                    options.append((team, "F"))

            elif player.position == "D":
                if can_add(team, "D"):
                    options.append((team, "D"))

            else:  # FLEX
                if can_add(team, "F"):
                    options.append((team, "F"))
                if can_add(team, "D"):
                    options.append((team, "D"))

        if not options:
            options = [
                (team_a, "F"),
                (team_b, "F"),
                (team_a, "D"),
                (team_b, "D")
            ]

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

def build_teams(roster_eml_path, player_data_xlsx_path):
    """
    Builds balanced teams by extracting player names from an EML roster file
    and looking up their rank/position from an Excel player data file.
    """
    # 2. Read player data (ranks and positions) from the Excel file
    df_players = pd.read_excel(player_data_xlsx_path)

    required_columns_xlsx = {"Name", "Rank", "Position"}
    # Ensure 'Name' column is treated as string to avoid issues with .strip().upper() later
    if 'Name' in df_players.columns:
        df_players['Name'] = df_players['Name'].astype(str)


    if not required_columns_xlsx.issubset(df_players.columns):
        raise ValueError(
            f"Player data Excel file must contain '{', '.join(required_columns_xlsx)}' columns."
        )

    # Create a lookup dictionary for player data for efficient access
    player_data_lookup = {
        row["Name"].strip().upper(): {
            "rank": int(cast(int, row["Rank"])),
            "position": cast(str, row["Position"]).strip().upper(),
        }
        for _, row in df_players.iterrows()
    }

    # 1. Get player names from EML roster file, validating against the Excel data
    eml_player_names = extract_players_from_eml(roster_eml_path, player_data_lookup)
    if not eml_player_names:
        raise ValueError("No player names extracted from the EML roster.")

    # 3. Construct Player objects for the game night roster
    players_for_tonight = []
    for name_from_eml in eml_player_names:
        normalized_name_eml = name_from_eml.strip().upper()
        if normalized_name_eml in player_data_lookup:
            data = player_data_lookup[normalized_name_eml]
            players_for_tonight.append(
                Player(
                    name=name_from_eml, # Use original casing for display purposes
                    rank=data["rank"],
                    position=data["position"]
                )
            )
        else:
            raise ValueError(
                f"Player '{name_from_eml}' from EML roster was not found in the player data Excel file."
            )

    if len(players_for_tonight) < 10:
        raise ValueError("Not enough selected players for team balancing (minimum 10 needed).")

    best = (None, None)
    best_diff = float("inf")

    for _ in range(ITERATIONS):
        # Pass the list of players constructed from EML names and XLSX data
        a, b = attempt_build(players_for_tonight.copy())
        diff = abs(a["total"] - b["total"])

        if diff < best_diff:
            best_diff = diff
            best = (deepcopy(a), deepcopy(b))

        if best_diff == 0:
            break

    return best[0], best[1], best_diff


# -----------------------------
# Export Workbook
# -----------------------------

def export_workbook(team_a, team_b, diff):

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
            "Light Team Total Rank",
            "Dark Team Total Rank",
            "Skill Difference"
        ],
        "Value": [
            team_a["total"],
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

def update_generate_button_state():
    """Enables the generate button only when both EML and XLSX files are selected."""
    eml_selected = eml_file_label.cget("text") != "No EML file selected"
    xlsx_selected = player_data_file_label.cget("text") != "No player data file selected"
    if eml_selected and xlsx_selected:
        generate_button.config(state="normal")
    else:
        generate_button.config(state="disabled")

def select_eml_file():
    """Opens a file dialog for selecting the EML roster file."""
    file_path = filedialog.askopenfilename(
        title="Select Roster EML File",
        filetypes=[("EML Files", "*.eml")]
    )
    if file_path:
        eml_file_label.config(text=file_path)
        update_generate_button_state()

def select_player_data_file():
    """Opens a file dialog for selecting the player data Excel file."""
    file_path = filedialog.askopenfilename(
        title="Select Player Data Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if file_path:
        player_data_file_label.config(text=file_path)
        update_generate_button_state()


def generate_teams():
    """
    Triggers the team generation process using the selected EML and Excel files.
    Handles errors and displays success/failure messages.
    """
    try:
        eml_file_path = eml_file_label.cget("text")
        player_data_xlsx_path = player_data_file_label.cget("text")

        if eml_file_path == "No EML file selected" or player_data_xlsx_path == "No player data file selected":
            raise ValueError("Please select both the EML roster file and the player data Excel file.")

        team_a, team_b, diff = build_teams(eml_file_path, player_data_xlsx_path)
        export_workbook(team_a, team_b, diff)

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
# Increase window size to accommodate new elements
root.geometry("460x400") # Original was 460x300
root.resizable(False, False)

title = tk.Label(root, text="🏒 Hockey Team Balancer", font=("Arial", 16, "bold"))
title.pack(pady=10)

# EML File Selection Components
select_eml_button = tk.Button(root, text="Select Roster EML File", command=select_eml_file, width=30)
select_eml_button.pack(pady=5)
eml_file_label = tk.Label(root, text="No EML file selected", wraplength=420)
eml_file_label.pack(pady=2)

# Player Data XLSX File Selection Components
select_player_data_button = tk.Button(root, text="Select Player Data Excel File", command=select_player_data_file, width=30)
select_player_data_button.pack(pady=5)
player_data_file_label = tk.Label(root, text="No player data file selected", wraplength=420)
player_data_file_label.pack(pady=2)

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
