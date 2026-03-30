# hockey-team-selection
Hockey Team Selection Code

This code requires the following setup on your machine to execute

    Python3

    Python libraries

        "pandas"
        "openpyxl"

You would also need a veriosn of Excel (or OpenOffice/Libre
Office) installed to be able to update the players as well as read the
results file.

This program will read an excel file of the following format:

    3 Columns from left to right

        Name - Players Name
        Rank - Value between 0 - 100
        Position - "F", "D" or "F/D"

The program also needs an email file to read the current weeks roster.
The email will ignore all lines above a line that starts with:

    "Which leaves you with a roster of"

The email should contain a list of player names (one per line) after that line.
Blank lines will be ignored.


The goal of the program is to select two teams "Dark" and "Light"
balancing the teams on both the Forward and Defence position based on
the player ratings

In order to use the program you must have selected a minimum of 10 PLayers to be playing

The program will output an excel workbook with three tabs

    "Light Team"
    "Dark Team"
    "Summary"
