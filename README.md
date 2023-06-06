# AUTOMATED PDF CONVERTER

## OVERVIEW
This program was created to automate the process of converting thousands of trial balance sheets on _Yardi Property Management_ to PDF format.

There are two programs that can be used to convert the pdfs. __automateSaves.py__ and __automateSaves1.0.py__ use different methods to achieve the same results. __automateSaves1.0.py__ is the quicker method for our purposes. I've included both of them because they are both viable options, and provide deep insights to automation with pyautogui and understanding decisions I made.

## TABLE OF CONTENTS
- UNDERSTANDING automateSaves.py
    - How it works
    - Running it
    - Stopping it
    - Considerations when recalibrating
- UNDERSTANDING automateSaves1.0.py
    - How it works
    - Running it
    - Stopping it
    - Considerations when recalibrating

## UNDERSTANDING automateSaves.py
### How it works
 The program is made for a three screen setup. One screen should have your Excel file with the file names. One screen should have _Yardi Property Management_ open. One screen should have this project's code open in whatever IDE you are using (you need to make sure that the IDE is not maximized/in full screen mode. Have it just float somewhere on screen, away from the top right corner of the screen). 

 It works as follows:
 1. Print a file from _Yardi Property Management_ to PDF format 
 2. Copy the corresponding file name from _Microsoft Excel_ 
 3. Save the file with the copied file name
 4. Close the file and any _Microsoft Word_ documents that may have popped out
 5. Repeat

 Using pyautogui and time modules, the program automates the monotonous task of retrieving PDF version of numerous files from ~~an outdated program that is~~ _Yardi Property Management_.

### Running it
 In order to run the program on your own machine, you must recalibrate all the coordinates to match your machine. Use __getPosition.py__ to retrieve them. You can keep a record of the coordinates in the __yardiPositions.md__ file.

The following conditions must be met:
1. pyautogui module is installed
2. _Excel_ window contains proper information
    - Run _SQL_ query to get all properties and their information and then order the properties by SCODE
    - Ensure that the order that _Yardi Property Management_ fetches data is the same as the order of your _SQL_ query
    (do a few test runs first)
3. Data in _Excel_ starts at 2nd row
3. Cell above data is selected on _Excel_
2. _Microsoft Word_ documents pop up on the third screen (The one with the IDE)
    - To do ensure this, click on the _Microsoft Word_ icon in the top right corner of the _Yardi Property Management_ window and drag the pop-up to the third screen and maximize it. _Yardi Property Management_ will remember this and open word documents the same way from now on.
4. Correct folder is selected when saving
    - To ensure this, navigate to the correct folder when _Yardi Property Management_ save window pops-up. _Yardi Property Management_ will remember this folder path for all future saves.
3. The print icon has not been selected before in the _Yardi Property Management_ window (you can't click on it twice)

To run it, run __automateSaves.py__ and let the program do the rest. 

### Stopping it
If you want to stop the program at anytime, move your mouse to the very top left corner of your leftmost screen (where coordinates are (0,0)) and leave it there until the next pyautogui call is made. The program will throw a FailSafeException.

### Considerations when recalibrating
To truly automate a process you need to ensure that an error will not occur or that they can be handled in a way that does not interfere with the rest of the process. This program relies on the former. Therefore, here are some possible reasons for which the process may act up if you recalibrate it.

1. Registering a click before arriving at the coordinate.
    - Having a moveTO and a click function executing one right after the other can cause such a problem. Put a pause in between the functions or find an alternative/simpler solution to the problem (e.g. hitting "enter"). 
2. Other program requires more time
    - Consider the time it takes for pop-up windows to open up. This program does not know how to check if a pop-up window is open.
3. Test runs with long pauses
    - Yes test it out, we know. But having test runs with long pauses can help identify two things:
        1. Where errors come from 
        2. How errors arise (If no errors during test runs with long pauses, maybe it is the lack of pause length that is the cause of the errors)
4. pyautogui does not do well with taskbar
    - I could not get it to click on the correct app on the taskbar. Maybe you can though.

Understanding where errors were coming from was the most challenging part of creating this program. Remember to keep your mind open, there are many paths that lead to the same result.

## UNDERSTANDING __automateSaves1.0.py__
### How it works
Long story short there was a quicker way to print the files to pdf. All you need to do is:
1. Copy file name from _Excel_
2. Paste file name in _Yardi Property Management_ pop-up window
3. Save file

### Running it
 In order to run the program on your own machine, you must recalibrate all the coordinates to match your machine. Use __getPosition.py__ to retrieve them. You can keep a record of the coordinates in the __yardiPositions.md__ file.

 The following conditions must be met:
1. pyautogui module is installed
2. _Excel_ window contains proper information in proper order
    - Run _SQL_ query to get all properties and their information and then order the properties by SCODE
    - Ensure that the order that _Yardi Property Management_ fetches data is the same as the order of your _SQL_ query
    (do a few test runs first)
3. Data in _Excel_ starts at 2nd row
3. Cell above data is selected in _Excel_
4. Correct folder is selected when saving
    - To ensure this, navigate to the correct folder when _Yardi Property Management_ save window pops-up. _Yardi Property Management_ will remember this folder path for all future saves.

To run it, run __automateSaves1.0.py__ and let the program do the rest. 

### Stopping it
If you want to stop the program at anytime, move your mouse to the very tope left corner of your leftmost screen (where coordinates are (0,0)) and leave it there until the next pyautogui call is made. The program will throw a FailSafeException.

### Considerations when recalibrating
Same considerations as __automateSaves.py__.

Most importantly, you ***MUST*** consider the following for this program:
1. Clipboard errors
    - The reason for the clipboard error is not 100% clear. What I believe happens is that the clipboard gets overcrowded and _Microsoft Office_ or the OS (idk which one is in charge of the clipboard) does not pop and push from the clipboard stack quick enough once it gets too full.
    - Solution that worked for me is as follows: 
        1. run "cmd /c echo.|clip" in the command line to reset the clipboard
        2. restart your computer
    - Clipboard errors have occured within the first 50 conversions and within the last 50 conversions so be aware of that.
2. Other program requires more time
    - Consider the time it takes for pop-up windows to open up.
    - For this program, if you click on the excel file before the pop-up window is opened in _Yardi Property Management_, then the pop-up window opens in excel. This would break the program. Make sure that the time between every iteration of the for loop is greater than the time it takes for the pop-up window to open. From my tests around 1.5 seconds works fine.
