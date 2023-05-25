# How to send report cards
This document is for teachers who teach any GL and/or VN class.

### I. Accessing GLVN spreadsheets

1. Enter https://drive.google.com/ into the address bar of your browser
2. Select `Shared with me`
3. If you can not see your class folder, you can search for your class folder in the search box.

<img width="588" alt="image" src="https://github.com/hungple/GLVN-database/assets/25112201/7520c410-6e03-41da-8e30-f45a9d8a8c11">


### II. Directory overview
In this section, GL1A is used as your class.

#### GL1A-Report-Cards: 
A directory to hold report cards. When you generate and email resport cards, the report cards will be generated in this folder. You don't need to save report cards if you send report cards to the parents. In case you need to save the report cards, please save them in your own Google drive or somewhere. The old report cards will be deleted after the school year ends.

#### GL1A:
A spreadsheet to hold your student information.

### III. Spreadshet overview

#### contacts: 
Contains contact information for your students. This sheet is read only. If you want to change anything, please contact your school administrators.

#### attendance-HK1 / attendance-HK2: 
You can use this sheet to keep track of your student attendance. However, this sheet is option. You are not required to use this sheet. You can clone this sheet to use for different purposes. You can also print out as blank sheet to use as offline sheet in class.

#### grades: 
This is the most important sheet of the spreadsheet. You are required to enter grades/scores and all other needed information of students during the school year.

<img width="740" alt="image" src="https://github.com/hungple/GLVN-database/assets/25112201/75e6e229-2f47-47bc-952d-5e82e66c9e28">

#### honor-roll: 
Lists students in order from the top score to lowest score. As a result, the top students will be sent to the honor roll. In most of the cases, you are not required to change anything unless you want different students to receive different awards. If you want to change anything in this sheet, please let the director of your school know. 

<img width="315" alt="image" src="https://github.com/hungple/GLVN-database/assets/25112201/cc483357-a75f-4f89-905f-8dbcff3f238b">

#### comment-review:
Lists students in order from the top score to lowest score. As a result, you can review your comments easier.

<img width="929" alt="image" src="https://github.com/hungple/GLVN-database/assets/25112201/1bc8d6c0-d40f-4219-9c59-f691cef9e2ba">



### IV. Entering grade points, attendance and comments
 

#### Part1 / Part2
This is participate column. You can use this column for participation in class, prayer and/or attendance. This column is required. The valid value is from 0 to 5. The script that generates report cards won't process if the column is not filled in.

#### HWrk1 / HWrk2
This is homework column. This column is required. The valid value is from 0 to 5. The script that generates report cards won't process if the column is not filled in.

#### Quiz1 / Quiz2
This is quiz column. This column is required. The valid value is from 0 to 15. The script that generates report cards won't process if the column is not filled in.

#### Exam1 / Exam2
This is exam column. This column is required. The valid value is from 0 to 25. The script that generates report cards won't process if the column is not filled in.

#### Extra1 / Extra2
This is extra credit column. This column is optional. The valid value is blank or from 0 to 20. If you leave this column blank, the script that generates report cards won't print out anything. If you enter 0, the script will print out 0 for the extra credit.

#### Note:
Even though this column can accept a value up to 20, it does not mean that you have to use this column. In fact, this column should be used for special cases only. If you give too many points for many students, there will be a lot of students can have a total point greater than 100 points but only limited top students can receive award certificates. For example, in the screen shot below, there are 4 students who have total score from 100 to 101.5; however, these students do not receive any award certificates. The parents of these students might ask why their child get perfect score (above 100) and not receive any award.

<img width="209" alt="image" src="https://github.com/hungple/GLVN-database/assets/25112201/b420d9a4-88ed-467b-9bae-bb14c75f4bef">

#### Absence1 / Absence2
This is absence column. This column is optional. The valid value is from 0 to 20. If you leave this column blank, the script that generates report cards won't print out anything.

#### Comment1 / Comment2
This is comment column. This column is optional. However, you should write very meaningful messages to your students. If you leave blank, it will look very bad in the report cards.

#### Action
This column is used with the `GLVN` menu items 
<br>
<img width="410" alt="image" src="https://github.com/hungple/GLVN-database/assets/25112201/eb38f3b7-9caa-4363-a2db-7a89dd77424c">

Enter `x` for each row of students if you want to generate report card. The report cards can be found in `GL1A-Report-Cards` folder.
![image](https://github.com/hungple/GLVN-database/assets/25112201/80ff9c61-e269-49d6-8037-af353fa700b5)


Enter `e` for each row of students if you want to generate and email report card. The report cards can be found in `GL1A-Report-Cards` folder. The emails and report cards can be found in your `sent` folder.
<img width="879" alt="image" src="https://github.com/hungple/GLVN-database/assets/25112201/ca646bbb-a705-4883-800e-76af80e42434">

If you are able to run the script, a message `Running script` is poped up as shown below. The message will be disappeared when the script is completed.

<img width="477" alt="image" src="https://github.com/hungple/GLVN-database/assets/25112201/9f91cf7b-d88a-4c38-8e82-665ec105c5a5">

#### Note:
- Before you send report cards to the parents, you should use the option `x` to generate report cards and review them before sending them to the parents.
- If you received any bounced email, check with the parent to get the correct email. Please contact your school administrators to update the email address.
- For the first time or from time to time, Google might ask you to authorize the script. Please follow section V. to authorize the script.
 
### V. Authorizing Google apps script

From time to time, Google might ask you to authorize or trust the script that is used in GLVN spreadsheets. To authorize, please follow the steps below:

![image](https://github.com/hungple/GLVN-database/assets/25112201/143981ee-5a19-4584-874a-3cd437b9e7a7)
- Click on `Continue` button.

![image](https://github.com/hungple/GLVN-database/assets/25112201/7ba1149e-f9f5-47ef-9e6b-d6b2e273e7f9)
- Click on the Google account that you want to use to send out report cards.

![image](https://github.com/hungple/GLVN-database/assets/25112201/53f3651d-662f-4176-9d1b-b29009a9e409)
- Click on `Advanced` link

![image](https://github.com/hungple/GLVN-database/assets/25112201/c2dfdf98-3073-44e7-97b2-b25dd11b02a4)
- Click on `Go to GL1A-app (unsafe)` link

![image](https://github.com/hungple/GLVN-database/assets/25112201/f3faf444-b62a-421b-860a-973bd9f2f6e6)
- Click on `Allow` button

#### Note:
If there is no popup windows showing the script is running, you can re-select GLVN menu item again. If you authorize from the previous step successfully, it should not ask you to authorize the script again.

<br>
or watch the first half of this video clip https://www.youtube.com/watch?v=4sFTQ9UAtuo&ab_channel=SheetsNinja

