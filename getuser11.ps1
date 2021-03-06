# Get user account information in AD
# By Randy 12th Aug 2016
$objOU=[ADSI]"LDAP://Ou=SHJ,OU=China,OU=Users,OU=Root2,DC=vstage,DC=co" # Enter your AD OU here
$searcher=new-object directoryservices.directorysearcher($objOU)
$searcher.Filter="(&(objectclass=user))" # define object type e.g. you can change it to workstation or something else
$Searcher.SearchScope = "Subtree"
$users=$searcher.findall()

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $True
$workbook = $excel.Workbooks.add()
$workbook.WorkSheets.item(1).Name = "UsersInfo" # Define Excel Sheet name
$sheet = $workbook.WorkSheets.Item("UsersInfo") # Swith to the sheet
$sheet.cells.item(1,1) = "Name" # Define row name
$sheet.cells.item(1,2) = "Description"
$sheet.cells.item(1,3) = "Dep"
$sheet.cells.item(1,4) = "Tel"
$sheet.cells.item(1,5) = "Win200Name"
$sheet.cells.item(1,6) = "Last Logon time"
$sheet.cells.item(1,7) = "ProfilePath"
$intRow=2

foreach ( $user in $users){
$Path=$user.Path # Get object path
$account=[ADSI]$Path # Get object by using ADSI
$sheet.cells.item($intRow,1) = $account.Name.Value
$sheet.cells.item($intRow,2) = $account.description.Value
$sheet.cells.item($intRow,3) = $account.physicaldeliveryofficename.Value
$sheet.cells.item($intRow,4) = $account.telephonenumber.Value
$sheet.cells.item($intRow,5) = $account.sAMAccountName.Value
$sheet.cells.item($intRow,6) = [datetime]::fromfiletime($account.ConvertLargeIntegerToInt64($account.lastlogontimestamp[0])) # Get lastlogontimestamp and convert to date format
$sheet.cells.item($intRow,7) = $account.Path
$intRow++
}

$strPath = ".\Test.xls" # Save Excel to your "My document" folder and name it as Test.xls
IF(Test-Path $strPath)
  { 
   Remove-Item $strPath
   $Excel.ActiveWorkbook.SaveAs($strPath)
  }
ELSE
  {
   $Excel.ActiveWorkbook.SaveAs($strPath)
  }
$Excel.Quit()