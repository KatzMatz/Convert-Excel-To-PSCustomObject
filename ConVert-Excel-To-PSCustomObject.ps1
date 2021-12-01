# Input  : excel book path, sheet name, range (, header)
# Output : PSCustomObject array

function Convert-Excel-To-PSCustomObject {
    param (
        [string] $XlsxPath,
        [string] $SheetName,
        [string] $Range,
        [string[]] $HeaderArray
    )
    
    # Input Validation
    if (!(Test-Path $XlsxPath)) {
        Write-Host ("Can't find {0}" -f $XlsxPath)
        return
    }

    # Header Config
    $Header = $false
    if ($null -ne $HeaderArray) {
        $Header = $true
    } else {
        $HeaderArray = @()
    }


    $excel = $null
    $book = $null
    $sheet = $null

    $resultArray = [PSCustomObject[]]@()

    try {
        # Wake up Excel
        $excel = New-Object -ComObject Excel.Application
        if ($null -eq $excel) {
            Write-Host "Can't open excel"
            return 
        }

        # Open the book
        $item = Get-Item $XlsxPath
        $UpdateLinks = 0 
        $ReadOnly = $True
        $Books = $excel.Workbooks
        $book = $Books.Open($item.FullName, $UpdateLinks, $ReadOnly)
        if ($null -eq $book) {
            Write-Host "Can't open the book"
            return
        }

        # Open the sheet
        $sheet = $book.Worksheets.item($SheetName)
        if ($null -eq $sheet) {
            Write-Host "Can't open the sheet"
            return
        }

        # Get Range cells
        $cells = $sheet.Range($Range)

        $numRow = $cells[$cells.count].Row
        $numColumn = $cells[$cells.count].Column


        if ($Header -and ($HeaderArray.count -ne $numColumn)) {
            Write-Host "Header array length is wrong"
            Write-Host ("HeaderArray Length is {0}, but column length of the range is {1}" -f $HeaderArray.Count, $numColumn)
            return
        } 

        # Write-Host ("({0}, {1})" -f $numRow, $numColumn)
        $numEmptyColumn = 0
        
        for ($r = 1; $r -le $numRow; $r = $r + 1) {
            # Define the object of the column
            $obj = [PSCustomObject]@{}

            for ($c = 1; $c -le $numColumn; $c = $c + 1) {
                $cellValue =  $cells[($r - 1)*$numColumn + $c].Text
                # Write-Host  ("{0} ({1}, {2})" -f  $cellValue, $r, $c)
                
                if ($null -eq $cellValue) {
                    $cellValue = ""
                }


                if (($Header -eq $false) -and ($r -eq 1)) {
                    if ($cellValue -eq "") {
                        $numEmptyColumn++
                        $cellValue = "(Empty Column " +$numEmptyColumn.ToString() +  ")"
                    }
                    $HeaderArray += $cellValue
                } else {
                    $obj | Add-Member -MemberType NoteProperty $HeaderArray[$c-1] -Value $cellValue
                }
            }
            if (($numRow -eq 1) -and ($Header -eq $false)) {
                $obj = [PSCustomObject]@{}
                foreach ($item in $HeaderArray) {
                    $obj | Add-Member -MemberType NoteProperty $item -Value ""
                }
            }
            $resultArray += $obj
        }

        
 
    } finally {
        # Write-Host "Close..."
        $sheet = $null
        $book = $null

        if ($null -ne $excel){
            $excel.Quit()
            $excel = $null
        }
    }

    return $resultArray

}
