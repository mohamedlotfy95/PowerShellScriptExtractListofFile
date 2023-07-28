# Set the path to the folder containing the files
$folderPath = "C:\Path\To\Folder"

# Set the path to the Excel file where the file names will be saved
$excelFilePath = "C:\Path\To\ExcelFile.xlsx"

# Get a list of the files in the specified folder
$fileList = Get-ChildItem $folderPath

# Create a new Excel COM object
$excelApp = New-Object -ComObject Excel.Application

# Make the Excel application visible
$excelApp.Visible = $true

# Add a new workbook to the Excel application
$workbook = $excelApp.Workbooks.Add()

# Get the first worksheet in the workbook
$worksheet = $workbook.Worksheets.Item(1)

# Set the value of the first cell in the worksheet to "File Name"
$worksheet.Cells.Item(1, 1).Value2 = "File Name"

# Iterate over the list of files
for ($fileIndex = 0; $fileIndex -lt $fileList.Count; $fileIndex++) {
    # Set the value of the cell in the next row to the name of the current file
    $worksheet.Cells.Item($fileIndex + 2, 1).Value2 = $fileList[$fileIndex].Name
}

# Save the workbook to the specified Excel file path
$workbook.SaveAs($excelFilePath)

# Close the workbook
$workbook.Close()

# Quit the Excel application
$excelApp.Quit()
