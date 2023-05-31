 $excelFile = $args[0]
 $o = $args[1]
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
