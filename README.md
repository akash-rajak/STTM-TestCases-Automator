## ✔ Open AI STTM TestCases Automation
- A python based tool to automate the test cases generation process for STTM Excel files using Open AI.
- In the model, first user need to select a STTM file from the File Selecting Dialog Box, and then the model will generate the test cases for the selected STTM file in different-different output presentation.
- And along with this output presentation, it also asks user if he/she wants to add those generated test cases on the ADO(Azure Devops) Dashboard.

****

### REQUIREMENTS :
- python 3
- os 
- tkinter 
- filedialog from tkinter
- pandas
- xlwt
- Path from pathlib
- openai
- docx
- json
- azure-devops
- Connection from azure.devops.connection
- BasicAuthentication from msrest.authentication
- JsonPatchOperation, WorkItem from azure.devops.v6_0.work_item_tracking.models

****

### How To Use it :
- User run the STTM_Automation.py file, on local system.
- After running this, user will be prompted to Enter Open AI Secret Key, ADO Dashboard details like Organization URL, Personal Access Token, Project Name and the Project ID(if user wants to add the test cases on the Azure Devops).
- After this user will be prompted to select the STTM Excel file using the pop-up file dialog box.
- After user has selected the file, in the backend the python model will follow the below steps inorder to generate the test cases automatically:
    - first searches for the STTM Sheet from the sheets in selected STTM Excel file.
    - then it generates the test cases using the Open AI API, in the following different format:
        - .xlsx
        - .docx
        - .txt
- Also if user chose to add the test cases on the ADO, then test cases will also be generated on Dashboard in the provide project_name and project id.


### Purpose :
- This model helps user to generate the test cases automatically without any human intervention.


### Compilation Steps :
- Install the mentioned required modules.
- After that download the code file, and run STTM_Automation.py on local system.
- Then the script will start running and user can explore it by selecting different and generating test cases for it.
