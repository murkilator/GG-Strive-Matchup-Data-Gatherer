import cv2
import numpy
import mss
import win32com.client
import time
import pyautogui
import xlsxwriter

# Generate an excel spreadsheet based on the results in result_matrix
def CreateExcel():
    workbook = xlsxwriter.Workbook("results.xlsx")

    # Create a worksheet to show win/loss raw data and one with win rates with formatted cells
    worksheet_raw = workbook.add_worksheet()
    worksheet_win_rate = workbook.add_worksheet()

    # Create formats with different colors for the win rate worksheet
    format_green = workbook.add_format({'num_format': '#,##0.00', 'bg_color' : '#7feb7f'})
    format_light_green = workbook.add_format({'num_format': '#,##0.00', 'bg_color' : '#c3f6c3'})
    format_white = workbook.add_format({'num_format': '#,##0.00', 'bg_color' : '#ffffff'})
    format_light_red = workbook.add_format({'num_format': '#,##0.00', 'bg_color' : '#f6c3c3'})
    format_red = workbook.add_format({'num_format': '#,##0.00', 'bg_color' : '#eb7f7f'})
    format_gray = workbook.add_format({'num_format': '#,##0.00', 'bg_color' : '#b4b4b4'})

    row = 1
    col = 1
    for row_character in characters:
        # Write character name row/column headers in both worksheets
        worksheet_raw.write(row, 0, row_character)
        worksheet_raw.write(0, col, row_character)
        worksheet_win_rate.write(row, 0, row_character)
        worksheet_win_rate.write(0, col, row_character)

        # For this row, go through all other characters and update the column value
        nested_col = 1
        for col_character in characters:
            row_character_wins = result_matrix[row_character][col_character]
            col_character_wins = result_matrix[col_character][row_character]
            total_matches = row_character_wins + col_character_wins

            format = format_white
            win_rate = "--"
            if total_matches > 0:
                win_rate = round(row_character_wins / total_matches, 2)

                if col_character == row_character:
                    format = format_gray
                else:
                    if (win_rate >= 0.54):
                        format = format_green
                    elif (win_rate >= 0.51):
                        format = format_light_green
                    elif (win_rate >= 0.49):
                        format = format_white
                    elif (win_rate >= 0.46):
                        format = format_light_red
                    else:
                        format = format_red

            worksheet_raw.write(row, nested_col, str(row_character_wins) + "-" + str(col_character_wins))
            worksheet_win_rate.write(row, nested_col, win_rate, format)
            nested_col = nested_col + 1

        row = row + 1
        col = col + 1

    workbook.close()

# Consider a max_val of 0.8 or better to be a match when searching for templates
# within images when using the cv2.TM_CCOEFF_NORMED method
MATCH_THRESHOLD = 0.8

# The maximum number of times in a row we can fail to identify any component of a
# replay result before giving up
MAX_FAIL_COUNT = 15

# The maximum number of replays to analyze
MAX_REPLAY_COUNT = 100000

# Amount of time to sleep before moving on to the next replay
SLEEP_TIME = 0.01

# Amount of time to sleep after failing to match any template before trying again
SLEEP_TIME_MATCH_FAIL = 0.1

# Amount of time to sleep after attempting to reset the position. This doesn't
# happen very often and it can take a while. Make it a long sleep to be safe
SLEEP_TIME_POSITION_RESET = 15

# The screen positions to grab when looking for player1, player2, and version info.
# Assumes 2560x1440 resolution
P1_SCREEN_POSITION      = {'top': 1085, 'left': 810,  'width': 220, 'height': 110}
P2_SCREEN_POSITION      = {'top': 1085, 'left': 1125, 'width': 220, 'height': 110}
VERSION_SCREEN_POSITION = {'top': 1090, 'left': 2030, 'width': 160, 'height': 70}

# List of characters. Character names must match the character template filenames
# in the templates folder
characters = ("sol", "ky", "may", "axl", "chipp", "pot", "faust", "millia", "zato", "ram", "leo", "nago", "gio", "anji", "ino", "gold", "jacko")

# Create a template for matching win/loss results and the game version
win_template = cv2.imread("templates\\win.png", cv2.IMREAD_GRAYSCALE)
lose_template = cv2.imread("templates\\lose.png", cv2.IMREAD_GRAYSCALE)
version_template = cv2.imread("templates\\version.png", cv2.IMREAD_GRAYSCALE)

# Create a map of character names and character templates and build the result
# matrix that the final results are stored in. The result matrix has keys of
# character names and values of dictionaries with keys of opponent character
# names and values of the win count
#
# Ex. {"sol" : {"ky" : 1, "gio" : 2}} shows that sol has 1 win against ky
# and 2 against gio
character_templates = {}
result_matrix = {}
for character in characters:
    character_templates[character] = cv2.imread("templates\\" + character + ".png", cv2.IMREAD_GRAYSCALE)

    results = {}
    for nested_character in characters:
        results[nested_character] = 0

    result_matrix[character] = results

# Force GGST to be in focus so it will take our key presses
wsh = win32com.client.Dispatch("WScript.Shell")
wsh.AppActivate("Guilty Gear -Strive-")

# Show the win/lose result on the replay screen
pyautogui.keyDown("n")

# Scroll down so the selected replay is fixed positionally on the screen
pyautogui.keyDown("s")
pyautogui.keyUp("s")
pyautogui.keyDown("s")
pyautogui.keyUp("s")
pyautogui.keyDown("s")
pyautogui.keyUp("s")
pyautogui.keyDown("s")
pyautogui.keyUp("s")
pyautogui.keyDown("s")
pyautogui.keyUp("s")
pyautogui.keyDown("s")
pyautogui.keyUp("s")

with mss.mss() as sct:
    fail_count = 0
    position_reset = False
    i = 0
    while i < MAX_REPLAY_COUNT:
        # Grab the sections of the screen that contain the player1 and player2 info
        # and convert them to pixel arrays
        p1_im = numpy.asarray(sct.grab(P1_SCREEN_POSITION))
        p2_im = numpy.asarray(sct.grab(P2_SCREEN_POSITION))

        # Convert the pixel arrays to gray so they can be matched with the gray
        # templates
        p1_im = cv2.cvtColor(p1_im, cv2.COLOR_BGR2GRAY)
        p2_im = cv2.cvtColor(p2_im, cv2.COLOR_BGR2GRAY)

        # Figure out which character is player1 and which character is player2 by
        # matching character templates
        p1_found = False
        p2_found = False
        p1_name = ""
        p2_name = ""
        for key, template in character_templates.items():
            if not p1_found:
                p1_result = cv2.matchTemplate(p1_im, template, cv2.TM_CCOEFF_NORMED)
                p1_min_val, p1_max_val, p1_min_loc, p1_max_loc = cv2.minMaxLoc(p1_result)

                # We found player1
                if p1_max_val >= MATCH_THRESHOLD:
                    p1_found = True
                    p1_name = key

            if not p2_found:
                p2_result = cv2.matchTemplate(p2_im, template, cv2.TM_CCOEFF_NORMED)
                p2_min_val, p2_max_val, p2_min_loc, p2_max_loc = cv2.minMaxLoc(p2_result)

                # We found player2
                if p2_max_val >= MATCH_THRESHOLD:
                    p2_found = True
                    p2_name = key
            
            if p1_found and p2_found:
                break

        # We didn't match player1 vs a template; increment fail_count and try again
        if not p1_found:
            fail_count = fail_count + 1

            # We failed as many times as allowed. It's possible the selected replay position
            # got reset to the top entry. If that's the case and we scroll up we'll be back
            # at the bottom of the list. If we haven't already tried this then reset fail_count,
            # scroll up, and try again. If we've already tried this and we are still failing then
            # abort
            if (fail_count == MAX_FAIL_COUNT):
                if position_reset:
                    print("Couldn't find player1 and position has already been reset. Aborting")
                    break
                else:
                    print("Couldn't find player1. Attempting to reset position")
                    position_reset = True
                    fail_count = 0
                    pyautogui.keyDown("w")
                    pyautogui.keyUp("w")
                    time.sleep(SLEEP_TIME_POSITION_RESET)
                    continue

            time.sleep(SLEEP_TIME_MATCH_FAIL)
            continue

        # We didn't match player2 vs a template; increment fail_count and try again
        if not p2_found:
            fail_count = fail_count + 1

            if (fail_count == MAX_FAIL_COUNT):
                if position_reset:
                    print("Couldn't find player2 and position has already been reset. Aborting")
                    break
                else:
                    print("Couldn't find player1. Attempting to reset position")
                    position_reset = True
                    fail_count = 0
                    pyautogui.keyDown("w")
                    pyautogui.keyUp("w")
                    time.sleep(SLEEP_TIME_POSITION_RESET)
                    continue

            time.sleep(SLEEP_TIME_MATCH_FAIL)
            continue

        # Figure out if player1 or player2 won by first checking if the win
        # template is contained in the player1 image and second checking if
        # the lose template is contained in the player1 image
        p1_win = False
        result = cv2.matchTemplate(p1_im, win_template, cv2.TM_CCOEFF_NORMED)
        win_min_val, win_max_val, win_min_loc, win_max_loc = cv2.minMaxLoc(result)
        if win_max_val >= MATCH_THRESHOLD:
            p1_win = True
        else:
            result = cv2.matchTemplate(p1_im, lose_template, cv2.TM_CCOEFF_NORMED)
            lose_min_val, lose_max_val, lose_min_loc, lose_max_loc = cv2.minMaxLoc(result)
            if lose_max_val >= MATCH_THRESHOLD:
                p1_win = False
            else:
                fail_count = fail_count + 1
                if fail_count > MAX_FAIL_COUNT:
                    if position_reset:
                        print("Failed to find the winner and the position has already been reset. Aborting. win_max_val: " + str(win_max_val) + " lose_max_val: " + str(lose_max_val))
                        break
                    else:
                        print("Failed to find the winner. Attempting to reset position. win_max_val: " + str(win_max_val) + " lose_max_val: " + str(lose_max_val))
                        position_reset = True
                        fail_count = 0
                        pyautogui.keyDown("w")
                        pyautogui.keyUp("w")
                        time.sleep(SLEEP_TIME_POSITION_RESET)
                        continue

                time.sleep(SLEEP_TIME_MATCH_FAIL)
                continue

        # Grab the section of the screen containing the version and convert to gray
        version_im = numpy.asarray(sct.grab(VERSION_SCREEN_POSITION))
        version_im = cv2.cvtColor(version_im, cv2.COLOR_BGR2GRAY)

        # Check if the version of this replay matches the version template
        result = cv2.matchTemplate(version_im, version_template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val < MATCH_THRESHOLD:
            fail_count = fail_count + 1
            if fail_count == MAX_FAIL_COUNT:
                print("Version mismatch. max_val: " + str(max_val))
                if position_reset:
                    break
                else:
                    position_reset = True
                    fail_count = 0
                    pyautogui.keyDown("w")
                    pyautogui.keyUp("w")
                    time.sleep(SLEEP_TIME_POSITION_RESET)
                    continue
                    

            time.sleep(SLEEP_TIME_MATCH_FAIL)
            continue

        # We identified everything successfully; reset the fail count and the position reset bool
        fail_count = 0
        position_reset = False

        # Increment the appropriate win count within the result matrix
        if p1_win:
            result_matrix[p1_name][p2_name] = result_matrix[p1_name][p2_name] + 1
        else:
            result_matrix[p2_name][p1_name] = result_matrix[p2_name][p1_name] + 1

        # Move on to the next replay
        pyautogui.keyDown("s")
        pyautogui.keyUp("s")
        time.sleep(SLEEP_TIME)

        # Print the current value of i for every 1000 replays for some indicator
        # of how far along we are
        i = i + 1
        if i % 1000 == 0:
            print("i: " + str(i))

    # We're finished; make the excel spreadsheet
    print(i)
    CreateExcel()
