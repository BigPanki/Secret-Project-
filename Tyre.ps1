$punctureResistant = $args[0]
$priceCategory = $args[1]

$data = Import-Excel -Path "\\HU-BUDFNP02\Operativ_Shared\Sales részére\Missing Tyre\Gumiméret segédtábla_20220812.xlsx" -WorksheetName "Summer"
    
        [String[]]$dimension = Import-Excel -Path $excelFile -WorksheetName "AJANLAT_netto" -ImportColumns @($o) -StartRow 26 -EndRow 26 -NoHeader
        $dimension = $dimension.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")
        $dimension = $dimension -replace '\s+|-|/', ''
        $dimension = $dimension -replace "[^\s0-9rR]"
        if ($dimension[0].Length -gt 8)
         {
            $1 = [int]$dimension[0].Substring(0, 5)
            $2 = [int]$dimension.Substring(8, 5)
            if ($1 -lt $2) 
            {
                $dimension = $dimension.Substring(8, 8)
            }
            else 
            {
                $dimension = $dimension.Substring(0, 8)    
            }
        }
             
        

        # Filter the data to only include rows with the cleaned dimension and the selected price category
        $filteredData = $data | Where-Object { $_.'Summer Dimension' -like "$dimension*" -and $_."Price Cat $priceCategory" -ne 0 }
    
        if ($punctureResistant) {
            # Filter the data to only include puncture-resistant tires
            $filteredData = $filteredData | Where-Object { $_.'Summer Dimension' -like "*RF*" -or $_.'Summer Dimension' -like "*RFX*" }
        }
    
        if ($filteredData.Count -eq 0) {
            Write-Output "No tires found for dimension $dimension with a non-zero price in price category $priceCategory"
        }
        else {
            # Sort the filtered data by price in ascending order
            $filteredData = $filteredData | Sort-Object -Property "Price Cat $priceCategory"
    
            # Loop through each row in the sorted filtered data and output the tire information
            foreach ($row in $filteredData) {
                $code = $row.'Summer Dimension'
                $price = $row."Price Cat $priceCategory"
                Write-Output "Summer: $code price: $price"
            }
        }
        
        
    
        $data = Import-Excel -Path "\\HU-BUDFNP02\Operativ_Shared\Sales részére\Missing Tyre\Gumiméret segédtábla_20220812.xlsx" -WorksheetName "Winter"
        # Filter the data to only include rows with the cleaned dimension and the selected price category
        $filteredData = $data | Where-Object { $_.'Winter Dimension' -like "$dimension*" -and $_."Price Cat $priceCategory" -ne 0 }
    
        if ($punctureResistant) {
            # Filter the data to only include puncture-resistant tires
            $filteredData = $filteredData | Where-Object { $_.'Winter Dimension' -like "*RF*" -or $_.'Winter Dimension' -like "*RFX*" }
        }
    
        if ($filteredData.Count -eq 0) {
            Write-Output "No tires found for dimension $dimension with a non-zero price in price category $priceCategory"
        }
        else {
            # Sort the filtered data by price in ascending order
            $filteredData = $filteredData | Sort-Object -Property "Price Cat $priceCategory"
    
            # Loop through each row in the sorted filtered data and output the tire information
            foreach ($row in $filteredData) {
                $code = $row.'Winter Dimension'
                $price = $row."Price Cat $priceCategory"
                Write-Output "Winter: $code price: $price"
            }
        }
