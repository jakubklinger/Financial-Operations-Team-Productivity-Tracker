# Excel Productivity Tracker | VBA Automation & Dashboard

📌**Project Overview**

A prototype Excel productivity tracking solution designed around a financial operations use case.

💡**Business Problems**

<ul>
<li>Productivity was being tracked manually</li>
<li>Different work types have different complexity</li>
<li>Raw task counts aren't sufficient for fair productivity comparison</li>
<li>Managers need an easy way to monitor trends</li>
</ul>

📊**Solution**

These problems can be solved with an Excel tool that combines several functionalities with an easy-to-use interface.
The tool utilizes control buttons and VBA code as well as PivotTable/Chart functionality. Daily productivity calculations can provide a useful way for employees to assess their performance. Additionally, this tool can be utilized as a part of productivity assessments by analysts and managers. 

🎯 **Objectives**
<ul>
<li>Tracking daily productivity based on several different work types </li>
<li>Providing an easy-to-use tool for employees that can be used further by analysts and managers</li>
<li>Implementing simple VBA script to minimize manual input and automate recurring actions </li>
</ul>

✨ Key Features
<ul>
<li>Utilizing weighted sums for the calculations</li>
<li>Conditional formatting assigned to productivity values</li>
<li>Buttons with dedicated VBA scripts for adding work items and new dates as well as removing items</li>
<li>PivotTable and Pivot for visualizing day-by-day progress</li>
<li>Feedback worksheet for continuous improvement</li>
</ul>

📋 **Files**
Financial Operations Team Productivity Tracker Ver 1.0.xlsm

🗄️ **Worksheets**

| Worksheet | Description |
|---------|-------------|
| `Dashboard`| Stores number of tasks processed daily, AHT (Average Handle Time) and productivity results. Utilizes buttons for adding a new day, new work item and clearing cell content. |
| `Chart` | Contains PivotTable and PivotChart for easy visual comparison of day-by-day productivity. |
| `Feedback` | Provides a structured dropdown‑based form created for a continuous improvement based on the feedback received. |
| `License` | Describes MIT License details. |


🛠️ **Stack**

<ul>
<li>Microsoft Excel 365</li>
<li>VBA (Visual Basic for Applications)</li>
<li>PivotTables & PivotCharts</li>
<li>Conditional Formatting</li>
</ul>

📊 **Dashboard Overview**

<img width="1115" height="398" alt="image" src="https://github.com/user-attachments/assets/4b8f6b5c-3743-45c3-858a-db3b407b5001" />

🧮 **Formulas**

** Average Handle Time (AHT) Sum & Productivity Formulas**

AHT formulas are calculated based on a weighted sum. Some tasks are longer due to extra steps or complexity. I estimated the weights as below:

<img width="905" height="71" alt="image" src="https://github.com/user-attachments/assets/20438cbd-5fa7-4658-a98b-b0b93ccd3178" />

When calculating the productivity, we need to remember that an employee won't be spending 100% of their login time on processing tasks. Lunch breaks, meetings, trainings and scheduled system updates reduce the available processing time. For this prototype, I assumed that approximately 400 of 480 working minutes are available for task processing. This assumption can be adjusted depending on the team's operating model.

Assuming that 400 out of 480 work minutes are spent on actual task processing, I based the formula on a weighted sum.

**AHT Formula** : <h3> **=SUMPRODUCT(B4:K4;$P$16:$Y$16)** </h3>

**Productivity Formula** : <h3> **=ROUND(L4/400; 2)** </h3> alternatively: **<h3> =ROUND((SUMPRODUCT(B4:K4;$P$16:$Y$16))/400; 2) </h3>**

** Average AHT and Average Productivity **

Calculating the average values provides employees with a clear view of their progress, especially when paired with conditional formatting (low results- red, average results- yellow, good results- green).

**Average AHT Formula** : **<h3> =IFERROR(ROUND(AVERAGEIF(L:L; ">0"); 2); 0) </h3>**
"IFERROR" operator was used to avoid blank rows throwing an error. Result is rounded to two decimal places.

**Average Productivity Formula** : **<h3> =IFERROR(ROUND(AVERAGEIF(M:M; ">0"); 2); 0) </h3>**

As above, "IFERROR" operator was used to avoid blank rows throwing an error and the result is rounded to two decimal places.

** Top Day (Best productivity) **
To find the highest-performance day, I used the MATCH operator to join the date (A column) of the highest value in M column (Productivity). 

**<h3> =TEXT(INDEX(A:A; MATCH(MAX(M:M); M:M; 0)); "DD/MM")
 & " - "
 & TEXT(MAX(M:M); "0,00%") </h3>**

**</>** **VBA Scripts **

💻 **Add Work Item**

This script allows users to easily add a new work item. The work type is selected from a dropdown list first and then button can be used to increase the value of the selected work type.

<img width="323" height="529" alt="image" src="https://github.com/user-attachments/assets/0bdc4806-65cc-4e3c-8e7f-0f4dfad59435" />


💻 **Delete Work Item**

This simple macro was included to allow quick clearing of the cells without disrupting the whole file.

<img width="531" height="110" alt="image" src="https://github.com/user-attachments/assets/cb6b7fca-dca9-4ec2-9b62-984b2567efe0" />


💻 **Add New Day**

Adding a new date to the Dashboard worksheet is easy with VBA usage. After clicking on the button, today's date is automatically added. What's more, the macro was created in a way that ensures that only the date and formulas are copied to the next day, while not duplicating the previous work amount values.

<img width="370" height="279" alt="image" src="https://github.com/user-attachments/assets/57750589-b7c8-493f-b06c-512f7bdaa4aa" />

💡 **Key Insights**

Good quality work tools need to be based on an approach that involves minimal user input while allowing for complex calculations. With this tool I wanted to showcase that utilizing simple VBA scripts can help with creating a file that is easy to use for the end user while still providing advanced functionality. Weighted productivity provides a more informative measure than task volume alone because different work types have different complexity levels.

✉ **---Contact me---**
For any questions, please contact me at jakub.klinger1996@gmail.com.
