# GLVN-database


### I. Authorizing Google apps script

The first time you run any menu item in GLVN menu, Google will ask you to authorize or trust the script that is used in GLVN spreadsheets. To authorize, please follow the instruction here https://kierandixon.com/authorize-google-apps-script/ or watch the first half of this video clip https://www.youtube.com/watch?v=4sFTQ9UAtuo&ab_channel=SheetsNinja

### II. Administrator's responsibilities

1. Help other teachers to use GLVN class spreadsheets. Administrators should not generate report cards for other teachers. If any teacher does not want to use Google spreadsheet, we can give him/her blank report cards to fill in manually. (easy)
 
2. Register new students during registration time or update student information such as updating email addresses. When teachers send out report cards, some email might be bounded due to invalid email addresses. (easy)
 
3. Share (or unshare) class spreadsheets, and 6 annual admin steps in students-master spreadsheet. The 6 annual admin steps need to be executed **only once** after we are done with the current school year and getting ready for new registration. (normal)
 
4. Test GLVN menu items in class spreadsheets and verify all functions. Administrators should test the class spreadsheets before teachers entering any data. If a bug is found after teachers entered data, we might need to spend more time to correct any issue. (normal)
 
5. Update code for spreadsheets students-master, students-extra, and class-library. Code in class-library is used for all class spreadsheets. However, in some cases, the library was not refreshed. In these cases, administrators will need to copy code for all class spreadsheets as well. Code can be found here https://github.com/hungple/GLVN-database. Please see section IV for how to copy the source code. (normal-hard)
 
6. Create custom GLxC or VNxC classes including GL9A (post confirmation). To create GLxC, clone from GL1A and update all ids. Similarly, to create VNxC, clone from VN1A. Please see section V for detail. (hard)
 
7. Become an owner of GLVN-database spreadsheets. (very hard. However, if you are computer engineer/programmer, the code is very easy to understand)
 

### III. Owner’s responsibilities:
 
1. Maintain source code in https://github.com/hungple/GLVN-database

2. Maintain all formulas, functions, and formats in all spreadsheets.


### IV. Copying source code to Google spreadsheets

Each GLVN spreadsheet has a release date in GLVN menu. If the release date is older than the one in the source code, you should copy the source code to the spreadsheet. For students-master and students-extra, just copy the respectively source code into the spreadsheet. For class spreadsheets (GL1A, VN1A..), you need to copy source code from class-library.gs into class-library spreadsheet. The code in class-library spreadsheet is used in all class spreadsheets. In some cases, the library does not work. In this case, you will need to copy code from class-library.gs to all class spreadsheets.

Here is the instruction how to copy source code to spreadsheets.
1. Select all the lines and then press Ctrl-C. Alternatively, you can click on the Raw button and then press Ctrl-A and then Ctrl-C to copy all the lines.
2. Open your class Google spreadsheet
3. Select grades tab > Extensions menu > Apps Script menu item
4. Click on the code editor
5. Press Ctrl-A and then Ctrl-V
6. Press Ctrl-S to save the new change
7. Close the code editor
8. Go back to your Google sheet and refresh it.


### V. Data flow
 
Here are flows of data between spreadsheets:
 
#### GL1A spreadsheet:
students-master.Std -> students-master.studentsclass -> students-master.GL1A -> GL1A.contacts -> GL1A.attendance-HK1/2, GL1A.graces
 
#### All students:
students-master.Std -> students-master.students -> students-extra.students-import ->students-extra.students-mini, students-extra.students-wide, students-extra.students-registration
 
#### Total points / final point:
GL1A.grades [column F] -> students-master.GL1A[column P] ->  students-master.Std [column AG and AH] (by selecting menu item Save student final points )
 
#### Honor roll:
GL1A.honor-roll -> students-extra.honor-gl-import ->  students-extra.honor-gl-1, honor-gl-2, honor-gl-3, students-extra.eucharist-certificates. The same applies to confirmation as well.


