# ==============================================================================
# สคริปต์สร้างไฟล์ Word ใบแจ้งผลการเรียน (ปพ.5) อัตโนมัติทุกระดับชั้น
# ตามแม่แบบ: แจ้งผลการเรียน69 ปรับใหม่ (ตราโรงเรียนบ้านสันดาบ + ตัวเลขอารบิก 100% + ผลประเมิน BSD-EVAL + คะแนนต่ำกว่า 25 สีแดง + หมายเหตุ)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$jsonPath = Join-Path $PSScriptRoot "data.json"
if (-not (Test-Path $jsonPath)) {
    Write-Error "ไม่พบไฟล์ data.json กรุณารันระบบดึงข้อมูลก่อน"
    exit
}

$templateFile = Join-Path $PSScriptRoot "แม่แบบ_แจ้งผลการเรียน69_ปรับใหม่.docx"
if (-not (Test-Path $templateFile)) {
    $srcPath = (Get-ChildItem -Path $PSScriptRoot -Filter "*ปรับใหม่*.docx" | Select-Object -First 1).FullName
    $src = [System.IO.File]::Open($srcPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $dst = [System.IO.File]::Create($templateFile)
    $src.CopyTo($dst)
    $src.Close()
    $dst.Close()
}

$outputBaseDir = Join-Path $PSScriptRoot "ใบแจ้งผลการเรียน_Word"
if (-not (Test-Path $outputBaseDir)) {
    New-Item -ItemType Directory -Path $outputBaseDir | Out-Null
}

$jsonText = [System.IO.File]::ReadAllText($jsonPath, [System.Text.Encoding]::UTF8)
$allClasses = $jsonText | ConvertFrom-Json
if ($allClasses.PSObject.Properties.Match('value').Count -gt 0) {
    $allClasses = $allClasses.value
}

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "  ระบบสร้างไฟล์ Word ใบแจ้งผลการเรียน โรงเรียนบ้านสันดาบ" -ForegroundColor Green
Write-Host "  (แม่แบบ: แจ้งผลการเรียน69 ปรับใหม่ + ตราโรงเรียน + เลขอารบิก + ผล BSD-EVAL)" -ForegroundColor Yellow
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "แม่แบบ: $(Split-Path $templateFile -Leaf)"
Write-Host "โฟลเดอร์ปลายทาง: $outputBaseDir`n"

function Set-CellText($cellNode, $text, $ns, $isRed = $false) {
    $p = $cellNode.SelectSingleNode(".//w:p", $ns)
    if (-not $p) { return }
    
    $runs = $p.SelectNodes("./w:r", $ns)
    $rPr = $null
    if ($runs.Count -gt 0) {
        $firstR = $runs.Item(0)
        $rPrNode = $firstR.SelectSingleNode("./w:rPr", $ns)
        if ($rPrNode) {
            $rPr = $rPrNode.CloneNode($true)
        }
    }
    if (-not $rPr) {
        $rPr = $cellNode.OwnerDocument.CreateElement("w", "rPr", $ns.LookupNamespace("w"))
    }
    
    if ($isRed) {
        $colorNode = $rPr.SelectSingleNode("./w:color", $ns)
        if (-not $colorNode) {
            $colorNode = $cellNode.OwnerDocument.CreateElement("w", "color", $ns.LookupNamespace("w"))
            [void]$rPr.AppendChild($colorNode)
        }
        [void]$colorNode.SetAttribute("val", $ns.LookupNamespace("w"), "FF0000")
        
        $bNode = $rPr.SelectSingleNode("./w:b", $ns)
        if (-not $bNode) {
            $bNode = $cellNode.OwnerDocument.CreateElement("w", "b", $ns.LookupNamespace("w"))
            [void]$rPr.AppendChild($bNode)
        }
    }
    
    # ลบเนื้อหาเดิมทั้งหมดในย่อหน้า
    while ($p.ChildNodes.Count -gt 0) {
        [void]$p.RemoveChild($p.ChildNodes.Item(0))
    }
    
    # สร้าง w:r และ w:t ใหม่
    $newR = $cellNode.OwnerDocument.CreateElement("w", "r", $ns.LookupNamespace("w"))
    if ($rPr) {
        [void]$newR.AppendChild($rPr)
    }
    $newT = $cellNode.OwnerDocument.CreateElement("w", "t", $ns.LookupNamespace("w"))
    $newT.InnerText = [string]$text
    [void]$newR.AppendChild($newT)
    [void]$p.AppendChild($newR)
}

function Format-EvalCheck($selectedRating) {
    $ratings = @("ดีเยี่ยม", "ดี", "ผ่าน", "ไม่ผ่าน")
    $out = @()
    foreach ($r in $ratings) {
        if ($r -eq $selectedRating) {
            $out += "☑ $r"
        } else {
            $out += "☐ $r"
        }
    }
    return ($out -join "   ")
}

function Generate-StudentDocx($cls, $student, $outPath) {
    Copy-Item -Path $templateFile -Destination $outPath -Force
    
    $zip = [System.IO.Compression.ZipFile]::Open($outPath, [System.IO.Compression.ZipArchiveMode]::Update)
    $entry = $zip.GetEntry('word/document.xml')
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    $xml = [xml]$reader.ReadToEnd()
    $reader.Close()
    $stream.Close()
    
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $ns.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
    $ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")

    # กำหนดขนาดตราโรงเรียน 0.8 นิ้ว (ความสูง 0.8 in = 731520 EMUs, กว้าง = 522514 EMUs)
    $extents = $xml.SelectNodes("//wp:extent", $ns)
    foreach ($ext in $extents) { $ext.cx = "522514"; $ext.cy = "731520" }
    $aExts = $xml.SelectNodes("//a:xfrm/a:ext", $ns)
    foreach ($ext in $aExts) { $ext.cx = "522514"; $ext.cy = "731520" }
    
    # 1. ปรับข้อความในย่อหน้า (Paragraphs)
    $paragraphs = $xml.SelectNodes("//w:p", $ns)
    $currentSection = ""

    foreach ($p in $paragraphs) {
        $fullTxt = ($p.SelectNodes(".//w:t", $ns) | ForEach-Object { $_.InnerText }) -join ""

        # ปรับระยะห่างบรรทัดชื่อด้านล่างให้กว้างขึ้น 6pt (120 twips)
        if ($fullTxt -match "ลงชื่อ" -or $fullTxt -match "ครูประจำชั้น" -or $fullTxt -match "หัวหน้างานบริหารวิชาการ" -or $fullTxt -match "ผู้อำนวยการ" -or $fullTxt -match "อนุมัติผลการเรียน") {
            $pPr = $p.SelectSingleNode("./w:pPr", $ns)
            if (-not $pPr) {
                $pPr = $xml.CreateElement("w", "pPr", $ns.LookupNamespace("w"))
                [void]$p.PrependChild($pPr)
            }
            $sp = $pPr.SelectSingleNode("./w:spacing", $ns)
            if (-not $sp) {
                $sp = $xml.CreateElement("w", "spacing", $ns.LookupNamespace("w"))
                [void]$pPr.AppendChild($sp)
            }
            [void]$sp.SetAttribute("after", $ns.LookupNamespace("w"), "120")
            [void]$sp.SetAttribute("line", $ns.LookupNamespace("w"), "280")
            [void]$sp.SetAttribute("lineRule", $ns.LookupNamespace("w"), "auto")
        }
        
        # ระดับชั้น และ ปีการศึกษา
        if ($fullTxt -match "ระดับชั้นประถมศึกษาปีที่.*ปีการศึกษา") {
            $tNodes = $p.SelectNodes(".//w:t", $ns)
            if ($tNodes.Count -gt 0) {
                $tNodes.Item(0).InnerText = "$($cls.ClassName)"
                for ($ti = 1; $ti -lt $tNodes.Count; $ti++) { $tNodes.Item($ti).InnerText = "" }
            }
        }
        # ข้อมูลนักเรียน (เลขที่ และ รหัสประจำตัว)
        elseif ($fullTxt -match "เด็กหญิง\.\.\.\.\.\.\.\.\." -or $fullTxt -match "\*\*\*\*\*\*\*\*\*\*\*\*\*\*" -or $fullTxt -match "เลขประจำตัว:") {
            $tNodes = $p.SelectNodes(".//w:t", $ns)
            if ($tNodes.Count -gt 0) {
                $tNodes.Item(0).InnerText = "เลขที่ $($student.No)   ชื่อ-นามสกุล: $($student.FullName)   เลขประจำตัว: $($student.StudentId)"
                for ($ti = 1; $ti -lt $tNodes.Count; $ti++) { $tNodes.Item($ti).InnerText = "" }
            }
        }
        # ติดตามหัวข้อผลการประเมิน
        elseif ($fullTxt -match "3\.\s*ผลการประเมินการอ่าน") {
            $currentSection = "reading"
        }
        elseif ($fullTxt -match "4\.\s*ผลการประเมินสมรรถนะ") {
            $currentSection = "competency"
        }
        elseif ($fullTxt -match "5\.\s*ผลการประเมินคุณลักษณะ") {
            $currentSection = "character"
        }
        elseif ($fullTxt -match "6\.\s*ความคิดเห็น") {
            $currentSection = "comment"
        }
        # ปรับเครื่องหมายผลการประเมิน ดีเยี่ยม ดี ผ่าน ไม่ผ่าน
        elseif ($fullTxt -match "ดีเยี่ยม.*ดี.*ผ่าน.*ไม่ผ่าน") {
            $rating = if ($currentSection -eq "reading") { $student.Evaluations.Reading }
                      elseif ($currentSection -eq "competency") { $student.Evaluations.Competency }
                      elseif ($currentSection -eq "character") { $student.Evaluations.Character }
                      else { "ดี" }
            
            $tNodes = $p.SelectNodes(".//w:t", $ns)
            if ($tNodes.Count -gt 0) {
                $tNodes.Item(0).InnerText = "  " + (Format-EvalCheck $rating)
                for ($ti = 1; $ti -lt $tNodes.Count; $ti++) { $tNodes.Item($ti).InnerText = "" }
            }
        }
        # ลายเซ็นครูประจำชั้น
        elseif ($fullTxt -match "ครูประจำชั้น" -and $fullTxt -match "\(\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\)") {
            $tNodes = $p.SelectNodes(".//w:t", $ns)
            foreach ($tn in $tNodes) {
                if ($tn.InnerText -match "\(\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\)") {
                    $tn.InnerText = "( $($cls.Teacher1) )"
                }
            }
        }
    }
    
    # 2. ตารางที่ 1: เวลาเรียน (เลขอารบิก)
    $tables = $xml.SelectNodes("//w:tbl", $ns)
    if ($tables.Count -ge 1) {
        $t1 = $tables.Item(0)
        $t1Rows = $t1.SelectNodes(".//w:tr", $ns)
        if ($t1Rows.Count -ge 2) {
            $r2Cells = $t1Rows.Item(1).SelectNodes(".//w:tc", $ns)
            if ($r2Cells.Count -ge 7) {
                Set-CellText $r2Cells.Item(1) "$($student.Attendance.Total)" $ns
                Set-CellText $r2Cells.Item(2) "$($student.Attendance.Present)" $ns
                Set-CellText $r2Cells.Item(3) "$($student.Attendance.Sick)" $ns
                Set-CellText $r2Cells.Item(4) "$($student.Attendance.Leave)" $ns
                Set-CellText $r2Cells.Item(5) "$($student.Attendance.Absent)" $ns
                Set-CellText $r2Cells.Item(6) "$($student.Attendance.Percent)" $ns
            }
        }
    }
    
    # 3. ตารางที่ 2: ผลสัมฤทธิ์ทางการเรียน
    if ($tables.Count -ge 2) {
        $t2 = $tables.Item(1)
        $t2Rows = $t2.SelectNodes(".//w:tr", $ns)
        
        # ปรับคะแนนเต็มแถวหัวข้อ
        if ($t2Rows.Count -ge 3) {
            $r3Cells = $t2Rows.Item(2).SelectNodes(".//w:tc", $ns)
            if ($r3Cells.Count -ge 4) {
                Set-CellText $r3Cells.Item(1) "(35 คะแนน)" $ns
                Set-CellText $r3Cells.Item(2) "(15 คะแนน)" $ns
                Set-CellText $r3Cells.Item(3) "(50 คะแนน)" $ns
            }
        }
        if ($t2Rows.Count -ge 13) {
            $r13Cells = $t2Rows.Item(12).SelectNodes(".//w:tc", $ns)
            if ($r13Cells.Count -ge 4) {
                Set-CellText $r13Cells.Item(1) "(40 คะแนน)" $ns
                Set-CellText $r13Cells.Item(2) "(10 คะแนน)" $ns
                Set-CellText $r13Cells.Item(3) "(50 คะแนน)" $ns
            }
        }
        
        foreach ($sub in $student.Scores) {
            $targetRowIdx = -1
            if ($sub.Index -ge 1 -and $sub.Index -le 9) {
                $targetRowIdx = 2 + $sub.Index
            } elseif ($sub.Index -ge 10 -and $sub.Index -le 15) {
                $targetRowIdx = 3 + $sub.Index
            }
            
            if ($targetRowIdx -ge 0 -and $targetRowIdx -lt $t2Rows.Count) {
                $rowCells = $t2Rows.Item($targetRowIdx).SelectNodes(".//w:tc", $ns)
                if ($rowCells.Count -ge 6) {
                    $totVal = 0
                    $isRed = $false
                    if ([double]::TryParse("$($sub.Total)", [ref]$totVal) -and $totVal -lt 25) {
                        $isRed = $true
                    }

                    Set-CellText $rowCells.Item(0) "$($sub.Index)" $ns
                    Set-CellText $rowCells.Item(1) "$($sub.Name)" $ns
                    Set-CellText $rowCells.Item(2) "$($sub.Code)" $ns
                    Set-CellText $rowCells.Item(3) "$($sub.Collect)" $ns
                    Set-CellText $rowCells.Item(4) "$($sub.Exam)" $ns
                    Set-CellText $rowCells.Item(5) "$($sub.Total)" $ns $isRed
                }
            }
        }
        
        # แถวเฉลี่ย (แถวสุดท้าย)
        $lastRow = $t2Rows.Item($t2Rows.Count - 1)
        $lastCells = $lastRow.SelectNodes(".//w:tc", $ns)
        if ($lastCells.Count -ge 2) {
            Set-CellText $lastCells.Item(1) "$($student.AverageScore)" $ns
        }

        # หมายเหตุท้ายตาราง: ติดตามแก้ไขผลการเรียน (เฉพาะเมื่อมีวิชาที่ได้คะแนนไม่ถึง 25 คะแนน)
        $hasUnder25 = $false
        foreach ($sub in $student.Scores) {
            $totVal = 0
            if ([double]::TryParse("$($sub.Total)", [ref]$totVal) -and $totVal -lt 25) {
                $hasUnder25 = $true
            }
        }

        $nextP = $t2.NextSibling
        if ($nextP -and $nextP.LocalName -eq "p") {
            while ($nextP.ChildNodes.Count -gt 0) { [void]$nextP.RemoveChild($nextP.ChildNodes.Item(0)) }
            if ($hasUnder25) {
                $r = $xml.CreateElement("w", "r", $ns.LookupNamespace("w"))
                $rPr = $xml.CreateElement("w", "rPr", $ns.LookupNamespace("w"))
                
                $rFonts = $xml.CreateElement("w", "rFonts", $ns.LookupNamespace("w"))
                [void]$rFonts.SetAttribute("cs", $ns.LookupNamespace("w"), "TH SarabunPSK")
                [void]$rFonts.SetAttribute("ascii", $ns.LookupNamespace("w"), "TH SarabunPSK")
                [void]$rPr.AppendChild($rFonts)
                
                $color = $xml.CreateElement("w", "color", $ns.LookupNamespace("w"))
                [void]$color.SetAttribute("val", $ns.LookupNamespace("w"), "FF0000")
                [void]$rPr.AppendChild($color)
                
                $b = $xml.CreateElement("w", "b", $ns.LookupNamespace("w"))
                [void]$rPr.AppendChild($b)
                
                $sz = $xml.CreateElement("w", "sz", $ns.LookupNamespace("w"))
                [void]$sz.SetAttribute("val", $ns.LookupNamespace("w"), "30") # 15pt
                [void]$rPr.AppendChild($sz)
                $szCs = $xml.CreateElement("w", "szCs", $ns.LookupNamespace("w"))
                [void]$szCs.SetAttribute("val", $ns.LookupNamespace("w"), "30")
                [void]$rPr.AppendChild($szCs)
                
                [void]$r.AppendChild($rPr)
                $t = $xml.CreateElement("w", "t", $ns.LookupNamespace("w"))
                $t.InnerText = "หมายเหตุ: ติดตามแก้ไขผลการเรียน"
                [void]$r.AppendChild($t)
                [void]$nextP.AppendChild($r)
            }
        }

        # ลบย่อหน้าว่างส่วนเกินหลังตารางที่ 2 เพื่อให้เอกสาร Word พอดี 1 หน้า A4 100%
        $sibling = if ($hasUnder25) { $nextP.NextSibling } else { $nextP }
        $blankRemoved = 0
        while ($sibling -and $sibling.LocalName -eq "p" -and $sibling.InnerText.Trim() -eq "" -and $blankRemoved -lt 4) {
            $toDel = $sibling
            $sibling = $sibling.NextSibling
            [void]$toDel.ParentNode.RemoveChild($toDel)
            $blankRemoved++
        }

        # ลบย่อหน้าว่างส่วนเกินท้ายเอกสารทั้งหมด ป้องกันไม่ให้ล้นไปหน้า 2
        $body = $xml.SelectSingleNode("//w:body", $ns)
        $allP = $body.SelectNodes("./w:p", $ns)
        for ($i = $allP.Count - 1; $i -ge 0; $i--) {
            $p = $allP.Item($i)
            $txt = ($p.SelectNodes(".//w:t", $ns) | ForEach-Object { $_.InnerText }) -join ""
            $drawings = $p.SelectNodes(".//w:drawing", $ns)
            if ([string]::IsNullOrWhiteSpace($txt) -and ($drawings -eq $null -or $drawings.Count -eq 0)) {
                [void]$body.RemoveChild($p)
            } else {
                break
            }
        }
    }
    
    # 4. บันทึกกลับ
    $entry.Delete()
    $newEntry = $zip.CreateEntry('word/document.xml')
    $writer = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
    $xml.Save($writer)
    $writer.Close()
    
    $zip.Dispose()
}

$totalGenerated = 0
foreach ($cls in $allClasses) {
    $shortName = if ($cls.ClassName -match 'ประถมศึกษาปีที่\s*([0-9]+)') { "ป." + $matches[1] } else { $cls.Id }
    $classDir = Join-Path $outputBaseDir $shortName
    if (-not (Test-Path $classDir)) {
        New-Item -ItemType Directory -Path $classDir | Out-Null
    }

    Write-Host "กำลังสร้างไฟล์สำหรับ $($cls.ClassName) ($($cls.Students.Count) คน)..." -ForegroundColor Yellow

    foreach ($student in $cls.Students) {
        $cleanFullName = ($student.FullName -replace '[\\/*?:"<>|]', '').Trim()
        $paddedNo = "{0:D2}" -f [int]$student.No
        $fileName = "ใบแจ้งผลการเรียน_${shortName}_เลขที่${paddedNo}_${cleanFullName}.docx"
        $filePath = Join-Path $classDir $fileName

        Generate-StudentDocx $cls $student $filePath
        $totalGenerated++
    }
    Write-Host "  -> สำเร็จ $($cls.Students.Count) ไฟล์ ในโฟลเดอร์ $shortName" -ForegroundColor Green
}

Write-Host "`n=================================================================" -ForegroundColor Cyan
Write-Host "  สร้างไฟล์ Word เสร็จสมบูรณ์ทั้งหมด $totalGenerated ไฟล์!" -ForegroundColor Green
Write-Host "  บันทึกไว้ที่: $outputBaseDir" -ForegroundColor White
Write-Host "=================================================================" -ForegroundColor Cyan
