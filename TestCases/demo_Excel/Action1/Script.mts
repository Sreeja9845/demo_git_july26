DataTable.AddSheet "sign_in_test_data_provided_by_sreeja"

 

'DataTable.ImportSheet"‪C:\Data\Navya_data.xlsx","login_data",3
DataTable.ImportSheet "C:\data\sreeja_data.xlsx","login_data",3


 

number_of_records=DataTable.GetSheet("sign_in_test_data_provided_by_sreeja").GetRowCount
msgbox number_of_records
For i=1 to number_of_records step 1

 

 

DataTable.getSheet(3).SetCurrentRow(i)
username=DataTable.Value("username",3)
password=DataTable.Value("password",3)
msgbox username
msgbox password

 

Browser("title:=Welcome: Mercury Tours").Page("title:=Welcome: Mercury Tours").WebEdit("name:=userName").Set username
Browser("title:=Welcome: Mercury Tours").Page("title:=Welcome: Mercury Tours").WebEdit("name:=password").SetSecure password
Browser("title:=Welcome: Mercury Tours").Page("title:=Welcome: Mercury Tours").WebButton("name:=submit").Click
Browser("title:=Login: Mercury Tours").Page("title:=Login: Mercury Tours").Link("name:=SIGN-OFF").Click
next

 

