# hockey-team-selection
Hockey Team Selection Code
This code requires the following setup on your machine to execute 

Python3
Python libraries - "pandas" and "openpyxl"

You would also need a veriosn of excel installed to be able to update the players as well as read the results file. 

This program will read an excel file of the following format: 
    4 Columns from left to right 
        Selected - Values Blank or "Yes" if selected to play that night
        Name - Players Name
        Rank - Value between 0 - 100
        Position - "F", "D" or "F/D"

The goal of the program is to select two teams "Dark" and "Light" balancing the teams on both the Forward and Defence position based on the player ratings

In order to use the program you must have selected a minimum of 10 PLayers to be playing

The program will output an excel workbook with three tabs
    "Light Team"
    "Dark Team" 
    "Summary"