Function ReadExcel($argPath ) {
    # Create excel Object
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $False
    $Book = $Excel.Workbooks.Open($argPath)
    $Sheet = $Excel.WorkSheets.Item(1)
    $Cells = $Sheet.Cells

    # Read Data 
    [Array]$Data = For ( $i = 2; $i -lt 200; $i++ )
    {
         [PSCustomObject]@{
            "Process Name" = $Cells.Item($i,1).Text
            "ID" = $Cells.Item($i,2).Text
            "CPU" = $Cells.Item($i,3).Text
            "VM" = $Cells.Item($i,4).Text
         }
         if( "" -eq $Cells.Item($i,1).Text ) {
             break
         }    
    }

    # Cleanup
    $Book.Close
    $Excel.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Excel) 

    # Output Data
    $Data | Format-Table -AutoSize
}

#--------------------------------
# Main
#--------------------------------
ReadExcel( "C:\work\Study\PowerShell\01_ProcessToExcel\output.xlsx") 
