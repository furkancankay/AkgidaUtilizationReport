# 📊 AkgidaUtilizationReport 📊
Thanks to this project, the efficiency, conditions and suitability of robots can be seen in the form of an Excel report. For the security of company data, I censored the database information.

## 🧬 Requirements
* `openpyxl   3.1.2`

## Our Goal:
Through this code, the data in the Excel file downloaded from the orchestrator platform within Ak Gıda RPA department is processed and a report is sent to RPA officials by calculating the utilization account.

## To Run
### 1. Data files
The Excel in the zip file downloaded via Orchestrator is extracted and its name is changed to data.csv.
* `degiskenlercsv_ortak_dosya_yolu = "C:\\Users\\furkan.cankaya\\Desktop\\ExcelKod\\data.csv"`

### 2. Input files
A csv file is created in the common area and the file path of this file is edited in the code.
* `degiskenlercsv_ortak_dosya_yolu = "C:\\Users\\furkan.cankaya\\Desktop\\ExcelKod\\variables.csv"`
* 
### 3. Set the paths
Insert the file path where you want to put the report into the code.
* `degiskenlercsv_ortak_dosya_yolu = "C:\\Users\\furkan.cankaya\\Desktop\\ExcelKod\\Utilizasyon.xlsx"`

##
Now, you are ready. When you run the code, it reads the data.csv data according to the input values ​​you entered, makes the utilization calculation and puts the utilization report in the folder you want.
##
