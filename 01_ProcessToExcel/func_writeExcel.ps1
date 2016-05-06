Function WriteExcel($argPath ) {
    # Create excel Object
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $False
    $Book = $Excel.Workbooks.Add()
    $Sheet = $Excel.WorkSheets.Item(1)
    $Cells = $Sheet.Cells
    
    # Write Header
    $Cells.Item(1,1) = "Process Name"
    $Cells.Item(1,2) = "ID"                       
    $Cells.Item(1,3) = "CPU"
    $Cells.Item(1,4) = "VM"

    # Write Body
    Get-Process | ForEach-Object{$i = 2} {
        $Cells.Item($i,1) = $_.ProcessName
        $Cells.Item($i,2) = $_.Id
        $Cells.Item($i,3) = $_.CPU
        $Cells.Item($i,4) = ($_.VM) / 1024 / 1024
        $i++
    }
    $Range = $Sheet.UsedRange
    $Range.EntireColumn.AutoFit() | Out-Null
    $Book.Charts.Add() | Out-Null

    # Save
    $Book.SaveAs($argPath)

    # Cleanup
    $Book.Close
    $Excel.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Excel)
}

#--------------------------------
# Main
#--------------------------------
WriteExcel( "C:\work\Study\PowerShell\01_ProcessToExcel\output.xlsx")                                                   C:\work\Study\PowerShell\01_ProcessToExcel

