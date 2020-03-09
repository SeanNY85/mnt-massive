Param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,
        [Parameter(Mandatory=$true)]
        [string]$DestPath
       )

#TESTING
$SourcePath="C:\Clients\Triangle\Source";
$Destpath="C:\Clients\Triangle\Destination"

#Continue with errors
$ErrorActionPreference= 'silentlycontinue'

Write "Loading Files... "
$files = Get-ChildItem -Path $SourcePath
Write "Loading Files $files."
Write "Files Loaded."
$output_type = "csv"

ForEach ($file in $files)
{
#Check for Worksheet named TestSheet
     $Excel = New-Object -ComObject Excel.Application
     $Excel.visible = $false
     $Excel.DisplayAlerts = $false
     $WorkBook = $Excel.Workbooks.Open($file.Fullname)
     $WorkSheets = $WorkBook.WorkSheets 

    if ($WorkBook.Worksheets.Count -gt 0) { 
        write-Output "Now processing: $WorkBook" 
        $FileFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV 

        $WorkBookFormat = $file -replace ".xlsx", ""
        $WorkBookName = $WorkBookFormat -replace " ","_"
        foreach($Worksheets in $WorkBook.Worksheets) {
            $ExtractedFileName = $WorkBookName+"."+$WorkSheets.Name + "." + $output_type 
            $WorkSheets.SaveAs($DestPath+"\"+$ExtractedFileName, $FileFormat) 

            write-Output "Created file: $ExtractedFileName"
        }
    } 

$WorkBook.Close() 
$Excel.Quit() 
Stop-Process -processname EXCEL
}
Read-host -prompt "The convetion has completed.  Press ENTER to close..."
clear-host;