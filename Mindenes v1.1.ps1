Add-Type -AssemblyName System.Windows.Forms

Add-Type -AssemblyName System.Windows.Forms
Import-Module ImportExcel
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern void mouse_event(int flags, int dx, int dy, int cButtons, int info);' -Name U32 -Namespace W;
Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class Win32 {
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
"@


$last_block = -1
function Set-CursorPosition 
{
    param(
        [Parameter(Mandatory = $true)]
        [int]$X,
        [Parameter(Mandatory = $true)]
        [int]$Y
    )

    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X, $Y)
}

function click 
{
    [W.U32]::mouse_event(6, 0, 0, 0, 0);
}

while ($true) {
    Write-Host "Please enter a block of code to execute (0-5) or 'q' to quit:
    0: Open excel file
    1: First Page
    2: Tyre
    3: Extra
    4: Extra total check
    5: NGM
    6: column chooser
    q: Quit"
    $input = Read-Host

    if ($input -eq 'q') {
        break
    }
    
    $block_num = [int]$input

    if ($block_num -lt 0 -or $block_num -gt 6) {
        Write-Host "Invalid input. Please enter a number between 1 and 6 or 'q' to quit."
        continue
    }

    # Execute the selected block of code
    if ($block_num -eq $last_block) {
        Write-Host "Executing block $block_num again..."
    }
    else {
        Write-Host "Executing block $block_num..."
    }

    #Excel open
    if ($block_num -eq 0) {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $dialog.Title = "Select an Excel File"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $excelFile = $dialog.FileName
        }
        
        
    }

    #First Page
    if ($block_num -eq 1) {
        [int]$o = Read-Host "Which column?"
        [String[]]$var = Import-Excel -Path "$excelFile" -WorksheetName "AJANLAT_brutto" -ImportColumns @($o) -StartRow 17 -EndRow 18 -NoHeader
        $var = $var.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")
        Set-Clipboard $var[0]
        
        # [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(915,360);
        Set-CursorPosition -X 915 -Y 360
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        Start-Sleep -Milliseconds 100
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        Start-Sleep -Milliseconds 100
            
        Set-Clipboard $var[1]
        # [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(673,360);
        Set-CursorPosition -X 673 -Y 360
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        Start-Sleep -Milliseconds 100
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        Start-Sleep -Milliseconds 100
        
        #Ár
        
        [String]$ar = Import-Excel -Path "$excelFile" -WorksheetName "AJANLAT_netto" -ImportColumns @($o) -StartRow 30 -EndRow 30 -NoHeader
        $ar = $ar.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")
        [Int]$ar = $ar
        $ar = [math]::Round($ar)
        
        Set-Clipboard $ar
        #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(262,360);
        Set-CursorPosition -X 262 -Y 360
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        Start-Sleep -Milliseconds 100
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        [System.Windows.Forms.SendKeys]::SendWait("^v")
            
        #Kedvezmény
        
        [String]$percent = Import-Excel -Path "$excelFile" -WorksheetName "AJANLAT_brutto" -ImportColumns @($o) -StartRow 37 -EndRow 37 -NoHeader
        $percent = $percent.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")
        [double]$percent = $percent
        [double]$percent = $percent
        $percent = $percent * 100
        $percent = [math]::Round($percent, 2)
        [string]$percent = $percent
        $asd = $percent.Replace(".", ",")
        $asd = $asd + "%"
        
        #ez itt^ megbaszhatja magát hogy így kell csinálni
        
        Set-Clipboard $asd
        #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(422,382);
        Set-CursorPosition -X 422 -Y 382
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        Start-Sleep -Milliseconds 100
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        Start-Sleep -Milliseconds 100
            
        #Delivery Cost
        
        [String]$deliv = Import-Excel -Path "$excelFile" -WorksheetName "AJANLAT_netto" -ImportColumns @($o) -StartRow 53 -EndRow 53 -NoHeader
        $deliv = $deliv.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")
        
        Set-Clipboard $deliv
        Start-Sleep -Milliseconds 100
        #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(689,384);
        Set-CursorPosition -X 689 -Y 384
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        Start-Sleep -Milliseconds 100
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        [System.Windows.Forms.SendKeys]::SendWait("^v")
            
        
        #Átadási idő
        
        [String]$comm = Import-Excel -Path "$excelFile" -WorksheetName "AJANLAT_brutto" -ImportColumns @($o) -StartRow 23 -EndRow 23 -NoHeader
        $comm = $comm.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")
        Set-Clipboard $comm
        #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(138,572);
        Set-CursorPosition -X 138 -Y 572
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        Start-Sleep -Milliseconds 100
        [W.U32]::mouse_event(6, 0, 0, 0, 0);
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        Start-Sleep -Milliseconds 100
        #print quot
        #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(128,640);
        Set-CursorPosition -X 128 -Y 640
        [W.U32]::mouse_event(6, 0, 0, 0, 0);

           
        $windowHandle = (get-process powershell_ise).MainWindowHandle
        [Win32]::SetForegroundWindow($windowHandle)
    } 
    elseif ($block_num -eq 2) {
        

        
        
       
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
        
        
        $windowHandle = (get-process powershell_ise).MainWindowHandle
        [Win32]::SetForegroundWindow($windowHandle)
        # Prompt the user for puncture resistance
        $punctureResistant = Read-Host "Do you need puncture resistant tires? (Y/N)"
        $punctureResistant = $punctureResistant -match "^[Yy]"
    
        # Prompt the user for the price category
        $priceCategory = Read-Host "Enter the price category (1, 2, or 3)"
        
        

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
        Write-Host""
        Write-Host""
        
    
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

       
        
        
    } 
    elseif ($block_num -eq 3) {
        #Utólag beszerelt
        #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(950,664);
        if ($quit -ne 'q') {
            [int]$o = Read-Host "Which column?"
            [String[]]$Extra = Import-Excel -Path "$excelFile" -WorksheetName "AJANLAT_netto" -ImportColumns @(1) -StartRow 58 -EndRow 76 -NoHeader
            $Extra = $Extra.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")
            [String[]]$Extraar = Import-Excel -Path "$excelFile" -WorksheetName "AJANLAT_netto" -ImportColumns @($o) -StartRow 58 -EndRow 76 -NoHeader
            $Extraar = $Extraar.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")
            [int[]]$Extraar = $Extraar

            #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(261, 242);
            Set-CursorPosition -X 261 -Y 242
            [int]$j = 0
            [int]$lineCounter = 0
            for ($i = 0; $i -lt $Extraar.Length; $i++) {
                if ($Extraar[$i] -ne "" -or $Extraar[$i] -eq "0") 
                {
                    try 
                    {
                        Set-Clipboard $Extra[$i]
                    }
                    catch 
                    {
                        break
                    }
                    
                    Start-Sleep -Milliseconds 100
                    #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(261,(242+$j));
                    Set-CursorPosition -X 261 -Y (242 + $j)
                    Start-Sleep -Milliseconds 100
                    [W.U32]::mouse_event(6, 0, 0, 0, 0);
                    Start-Sleep -Milliseconds 100
                    [W.U32]::mouse_event(6, 0, 0, 0, 0);
                    Start-Sleep -Milliseconds 100
                    [System.Windows.Forms.SendKeys]::SendWait("^v")
                    Start-Sleep -Milliseconds 100
                    $j += 20
                    $lineCounter++
                } 
                else {
                    break
                }
            }
            [int]$j = 0
            for ($i = 0; $i -lt $Extra.Length; $i++) {
                if ($Extraar[$i] -ne "") {
                    Set-Clipboard $Extraar[$i]
                    Start-Sleep -Milliseconds 100
                    #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(565,(242+$j));
                    Set-CursorPosition -X 565 -Y (242 + $j)
                    Start-Sleep -Milliseconds 100
                    [W.U32]::mouse_event(6, 0, 0, 0, 0);
                    Start-Sleep -Milliseconds 100
                    [W.U32]::mouse_event(6, 0, 0, 0, 0);
                    Start-Sleep -Milliseconds 100
                    [System.Windows.Forms.SendKeys]::SendWait("^v")
                    Start-Sleep -Milliseconds 100
                    $j += 20
                } 
                else {
                    break
                }
            }

            
        }
     
        Set-CursorPosition -X 968 -Y 565
        while ($lineCounter -gt 0) {
            
            Start-Sleep -Milliseconds 50
            [W.U32]::mouse_event(6, 0, 0, 0, 0);
            $lineCounter--
        }

        [String[]]$Extra = Import-Excel -Path "$excelFile" -WorksheetName "AJANLAT_brutto" -ImportColumns @(1) -StartRow 76 -EndRow 146 -NoHeader
        $Extra = $Extra.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")

        [String[]]$Extraar = Import-Excel -Path "$excelFile" -WorksheetName "AJANLAT_netto" -ImportColumns @($o) -StartRow 80 -EndRow 150 -NoHeader
        $Extraar = $Extraar.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")

        [int]$j = 0
        [int]$k = 0
        [int]$l = 0
        
        
        for ($i = 0; $i -lt $Extraar.Length; $i++) {
            if ($Extraar[$i] -ne "" -or $Extraar[$i] -eq "0") {
                Set-Clipboard $Extra[$i]
                Start-Sleep -Milliseconds 100
                Set-CursorPosition -X 261 -Y (242 + $j)
                Start-Sleep -Milliseconds 100
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
                Start-Sleep -Milliseconds 100
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
                Start-Sleep -Milliseconds 100
                [System.Windows.Forms.SendKeys]::SendWait("^v")
                Start-Sleep -Milliseconds 100
                $j += 20
                $k++
                if ($k % 17 -eq 0) {
                    $j = 0
                    $l++
                    Set-CursorPosition -X 968 -Y 565
                    for ($x = 0; $x -lt 17; $x++) {
                        Start-Sleep -Milliseconds 50
                        [W.U32]::mouse_event(6, 0, 0, 0, 0);
                    }
                    
                }
            }
                
        }
        if ($l -gt 0) {
            Set-CursorPosition -X 968 -Y 220
            for ($x = 0; $x -lt 17; $x++) {
                Start-Sleep -Milliseconds 50
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
            }
            
        }
        [int]$j = 0
        [int]$k = 0
        for ($i = 0; $i -lt $Extraar.Length; $i++) {
            if ($Extraar[$i] -ne "" -or $Extraar[$i] -eq "0") {
                $roundedExtra = [math]::Round($Extraar[$i], 0)
                Set-Clipboard $roundedExtra
                Start-Sleep -Milliseconds 100
                #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(565,(242+$j));
                Set-CursorPosition -X 565 -Y (242 + $j)
                Start-Sleep -Milliseconds 100
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
                Start-Sleep -Milliseconds 100
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
                Start-Sleep -Milliseconds 100
                [System.Windows.Forms.SendKeys]::SendWait("^v")
                Start-Sleep -Milliseconds 100
                $j += 20
                $k++
                if ($k % 17 -eq 0) {
                    $j = 0
                    #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(968,565);
                    Set-CursorPosition -X 968 -Y 565
                    for ($x = 0; $x -lt 17; $x++) {
                        Start-Sleep -Milliseconds 50
                        [W.U32]::mouse_event(6, 0, 0, 0, 0);
                    }
                }
            }
                
        }
            
        if ($k % 16 -eq 0) {
            #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(968,220);
            Set-CursorPosition -X 968 -Y 220
            for ($i = 2; $i -le $k; $i++) {
                Start-Sleep -Milliseconds 50
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
            }
            Start-Sleep -Milliseconds 50
        }

        Set-Clipboard $asd

        
    }
    elseif ($block_num -eq 4) {
        [String[]]$excelTotalExtra = Import-Excel -Path $excelFile -WorksheetName "AJANLAT_netto" -ImportColumns @($o) -StartRow 33 -EndRow 33 -NoHeader
        $excelTotalExtra = $excelTotalExtra.replace("@", "").replace("{", "").replace("P1", "").replace("=", "").replace("}", "")
        [double[]]$excelTotalExtra = $excelTotalExtra
        #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(628, 618);
        #Extras OK button
        Set-CursorPosition -X 955 -Y 669
            Start-Sleep -Milliseconds 100
            [W.U32]::mouse_event(6, 0, 0, 0, 0);
            
            Start-Sleep -Milliseconds 100
            $windowHandle = (get-process powershell_ise).MainWindowHandle
            [Win32]::SetForegroundWindow($windowHandle)
            Start-sleep -Milliseconds 100
            Read-Host 'Press enter to continue.'
            
            
        Set-CursorPosition -X 628 -Y 618
            Start-Sleep -Milliseconds 100
            [W.U32]::mouse_event(6, 0, 0, 0, 0);
            Start-Sleep -Milliseconds 100
            [W.U32]::mouse_event(6, 0, 0, 0, 0);
        
       

        [System.Windows.Forms.SendKeys]::SendWait("^c")
        Start-Sleep -Milliseconds 300  
        try 
        {
            [double]$totalExtra = [System.Windows.Clipboard]::GetText();
        }
        catch 
        {
            Write-Host "Null Value on clipboard!"
        }
        
        $totalExtra = $totalExtra / 100
        if ($totalExtra -ge $excelTotalExtra[0] - 10 -and $totalExtra -le $excelTotalExtra[0] + 10) 
        {  
            Write-Host "Total of extras MATCH"
            $windowHandle = (get-process powershell_ise).MainWindowHandle
            [Win32]::SetForegroundWindow($windowHandle)
        }
        else {
            Write-Host "Total of extras DOES NOT MATCH"
        }
    }
    elseif ($block_num -eq 5) {
        
            
            Set-CursorPosition -X 955 -Y 669
            Start-Sleep -Seconds 1
            [W.U32]::mouse_event(6, 0, 0, 0, 0);
            Start-Sleep -Seconds 2
            
            Set-CursorPosition -X 1387 -Y 834
            [W.U32]::mouse_event(6, 0, 0, 0, 0);
            Start-Sleep -Seconds 2
            
            $windowHandle = (get-process powershell_ise).MainWindowHandle
            [Win32]::SetForegroundWindow($windowHandle)
            try {
                #NGM számítás
                [double]$userNGM = Read-Host("What NGM do you want?")
            
    
                #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(252, 238);
                Set-CursorPosition -X 252 -Y 238
                Start-Sleep -Milliseconds 100
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
            
                Start-Sleep -Milliseconds 300
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
                Start-Sleep -Milliseconds 300
    
                [System.Windows.Forms.SendKeys]::SendWait("^c")
                Start-Sleep -Milliseconds 300
                
                [double]$NGM = [Windows.Forms.Clipboard]::GetText();
                $NGM = $NGM / 100
                $NGM = - $NGM + $userNGM
                Set-Clipboard $NGM
                Start-Sleep -Milliseconds 100
                #[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(252, 324);
                Set-CursorPosition -X 252 -Y 324
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
                Start-Sleep -Milliseconds 100
                [W.U32]::mouse_event(6, 0, 0, 0, 0);
                [System.Windows.Forms.SendKeys]::SendWait("^v") 
            }
            catch {
                Write-Host "Invalid NGM, exiting"
            }
        }
        

    }
    elseif ($block_num -eq 6) {
        [int]$o = Read-Host "Which Column?"
    }

     

    # Update the last_block variable
    $last_block = $block_num

