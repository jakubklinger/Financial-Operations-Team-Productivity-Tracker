# Financial-Operations-Team-Productivity-Tracker

📌 **Project Overview**



The aim was to create an Excel tracking tool that could be utilised by investment/mutual funds companies. 

It combines easy to use user interface with control buttons and VBA code as well as PivotTable/Chart functionality. Daily calculations of productivity could be an excellent solution for self assestment of employee's results. What's more this tool can be utilized as a part of the productivity assetments by analysts and managers. 

Feedback worksheet was added for a continous improvement.


🎯 **Objectives**
<ul>
<li>Tracking daily productivity based on several different worktypes with assigned weights</li>
<li>Providing easy to use tool for employees that can be used further by analysts and managers</li>
<li>Utilizing weiged sums for the calculations</li>
<li>Conditional formating assigned to productivity values</li>
<li>Implementing simple VBA script to minimalize manual input and automate reccuring actions (adding a line for new date and preparing charts)</li>
</ul>

📋 **Files**
Financial Operations Team Productivity Tracker Ver 1.0.xlsm

🗄️ **Worksheets**

| Worksheet | Description |
|---------|-------------|
| `Dashboard`| Stores number of tasks processed daily, AHT (Average Handle Time) and productivity results. Utilizes buttons for adding new day, new work item and clear cell's content. |
| `Chart` | Contains PivotTable and Pivot Chart for easy visual comparison of day by day productivity. |
| `Feedback` | Dropdown based sheet created for a continous improvement based on the feedback received. |
| `License` | Describes MIT License details. |


🛠️ **Stack**
Microsoft Excel 365

**📊Dashboard Overview**

<img width="1115" height="398" alt="image" src="https://github.com/user-attachments/assets/4b8f6b5c-3743-45c3-858a-db3b407b5001" />

**🧮Formulas**

** Average Handle Time (AHT) Sum & Procuctivity Formulas*

AHT formulas are calculated based on a weighted sum. Some tasks are longer due to extra steps or complexity. I estimated the weights as below:

<img width="905" height="71" alt="image" src="https://github.com/user-attachments/assets/20438cbd-5fa7-4658-a98b-b0b93ccd3178" />

When calcuating the productivity, we need to remember that employee won't be spending 100% of their login time on processing tasks. Lunch break, meetings, trainings or scheduled system updates will reduce raw processing time. We can assume that on average, 30-40 minutes a day are spent on the breaks and around 30-90 on meetings. Usually, daily meetings are kept short for crucial updates, taking between 15-30 mintes. Periodically,  meetings with Q&A sessions or integration focused events might be longer. 

Assuming that 400 out of 480 work minutes are spend on actual task processing, I based the formula on a weigthed sum.

**AHT Formula** : <h3> **=SUMPRODUCT(B4:K4;$P$16:$Y$16)** </h3>

**Productivity Formula** : <h3> **=ROUND(L4/400; 2)** </h3> alternatively: **<h3> =ROUND((SUMPRODUCT(B4:K4;$P$16:$Y$16))/400; 2) </h3> **

** Average AHT and Average Productivity **

Calculating the average values might give employees a data representation of their progress, especially when paired with conditional formatting (low results- red, average results- yellow, good results- green).

**Average AHT Formula** : **<h3> =IFERROR(ROUND(AVERAGEIF(L:L; ">0"); 2); 0) </h3>**
"IFERROR" operator was used to avoid blank rows throwing an error. Result is rounded up to two decimal places.

**Average Productivity Formula** : **<h3> =IFERROR(ROUND(AVERAGEIF(M:M; ">0"); 2); 0) </h3>**

As above, "IFERROR" operator was used to avoid blank rows throwing an error and the result is rounded up to two decimal places.

** Top Day (Best productivity) **
To find the top performance day, I used the MATCH operator to join the date (A column) of the highest valey in M column (Productivity). 

**<h3> =TEXT(INDEX(A:A; MATCH(MAX(M:M); M:M; 0)); "DD/MM")
 & " - "
 & TEXT(MAX(M:M); "0,00%") </h3>**

**</>** **VBA Scripts **

💻 **Add Work Item**

This script was used to allow easy adding of new work item. The worktype is selected from a dropdown list first and then button can be used to increase the value of selected work type.

<img width="323" height="529" alt="image" src="https://github.com/user-attachments/assets/0bdc4806-65cc-4e3c-8e7f-0f4dfad59435" />


💻 **Delete Work Item**

This simple macro was included to allow quick clearing of the cells without disrupting the whole file.

<img width="531" height="110" alt="image" src="https://github.com/user-attachments/assets/cb6b7fca-dca9-4ec2-9b62-984b2567efe0" />


💻 **Add New Day**

Adding a new date for the Data worksheet is easy with VBA usage. After clicking on the button, today's date is automatically added. What's more, macro was created in a way that ensures that only date and formulas are copied to next day, while not duplicating the previous work ammount values.

<img width="370" height="279" alt="image" src="https://github.com/user-attachments/assets/57750589-b7c8-493f-b06c-512f7bdaa4aa" />

💡 **Key Insight**

Good quality work tools need to be based on an approach that involves minimum user input while allowing for complex calculations. With this tool I wanted to showcase that utilizing simple VBA scripts can help with creating a file that is easy to use for an end user while still providing advanced functionality.


✉ **---Contact me---**
For any questions, please contact me at jakub.klinger1996@gmail.com.

