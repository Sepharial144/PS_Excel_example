# Path to file
$Path = "E:\PowerShell_My_Developments\PS_Excel" #файл Source

# Name of xlsx file
$File = "Example.xlsx"

# Make full path to file
$pathToFile = $Path + "\" + $File

# Create object Excel
$objExcel = New-Object -ComObject Excel.Application

# Make docment as visible 
$objExcel.Visible = $true

# Open file xlsx
$Workbook = $objExcel.Workbooks.Open($pathToFile)

# Choose first sheet
$Worksheet = $Workbook.Sheets.item(1)


# --------------Write into cells ----------------

# Cell A1
$Worksheet.cells.Item(1, 1) = "This A1 cell"

# Cell B1
$Worksheet.cells.Item(1, 2) = "This B1 cell"

# Cell A2
$Worksheet.cells.Item(2, 1) = "This A2 cell"

# Cell B2
$Worksheet.cells.Item(2, 2) = "This B2 cell"

# --------------Read from cells ----------------

# Find row count in range of table
$countRow = $Worksheet.UsedRange.Rows.count

# Display rows in file
Write-Host "Excel file have :$countRow rows" 

for($i = 1; $i -le $countRow; $i++) { 

    $var_1 = $Worksheet.cells.Item($i, 1).Text

    $var_2 = $Worksheet.cells.Item($i, 2).Text

    Write-Host "Value A$i :$var_1"
    Write-Host "Value B$i :$var_2"

}

# Save without popup
$Workbook.Application.DisplayAlerts = $false

# Save and close excel workbook
$Workbook.SaveAs($pathToFile)

# Return property as default param
$Workbook.Application.DisplayAlerts = $false

# Close application
$objExcel.Quit()

# Release COM - object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)

# Remove variable named as objExcel
Remove-Variable objExcel