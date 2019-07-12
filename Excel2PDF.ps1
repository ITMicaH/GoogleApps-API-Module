#$RootFolder = "C:\Users\micah\OneDrive\Documenten\Stand Leyenburg"
$RootFolder = "C:\Users\zoum2203\OneDrive\Documenten\Stand Leyenburg"
$ExcelFile = "$RootFolder\Stand Leyenburg.xlsm"
$PictureFolder = $RootFolder
$Months = 7#..9
$Day = 'Woensdag' # Woensdag Zaterdag

#region functions

function FitToPage
{
    try
    {
        $Document.FitToPages()
    }
    catch{
        switch -w ($_.exception)
        {
            '*unable to shrink*' {
                $WordPic.ScaleHeight = $WordPic.ScaleHeight - 10
                $WordPic.ScaleWidth = $WordPic.ScaleWidth - 10
                FitToPage
            }
            '*document is already*' {
                return
            }
            default {
                Write-Error $_
            }
        }
    }
}

#endregion functions

switch ($Day)
{
    Woensdag {
        $Shifts = @(
            '10:00'
            '11:30'
            '13:00'
        )
    }
    Zaterdag {
        $Shifts = @(
            '11:30'
            '13:00'
            '14:30'
        )
    }
}

#region Import Schedule from Excel

$Schema = Import-Excel $ExcelFile -WorksheetName "Schema $Day"
$ShiftSchedule = foreach ($Month in $Months)
{
    $Dates = $Schema.Where{$_.Datum.Month -eq $Month} | select * -ExcludeProperty Standwerker,Shift*,Totaal
    foreach ($Date in $Dates)
    {
        for ($i = 1; $i -le $Shifts.Count; $i++)
        { 
            $StartTime = $Shifts[$i-1]
            $Volunteers = ($Date | Get-Member -MemberType NoteProperty).where{
                $Date.($_.Name) -eq $i
            }.Name.foreach{
                [Volunteer]::new($_)
            }
            [pscustomobject]@{
                Start = $Date.Datum + [timespan]$StartTime
                End = $Date.Datum + [timespan]$StartTime + [timespan]'1:30'
                ShiftTime = "$StartTime - $(([timespan]$StartTime + [timespan]'1:30').ToString('hh\:mm'))"
                Volunteers = $Volunteers
            }
        }
    }
}

#endregion Import Schedule from Excel

#region Create Schedule in Word and PDF

$Word = New-Object -ComObject Word.Application
$Word.Visible = $true
$GroupByMonth = $ShiftSchedule | group {$_.Start.month}
foreach ($Group in $GroupByMonth)
{
    $TableContents = $Group.Group | group ShiftTime | ForEach{
        $HT = [ordered]@{
            Shift = $_.Name
        }
        $_.Group | sort Start | foreach{
            If ($_.Volunteers)
            {
                $HT.Add($_.Start.ToString('d MMMM'),($_.Volunteers -join "`v"))
            }
            else
            {
                $HT.Add($_.Start.ToString('d MMMM'),"-")
            }
        }
        [pscustomobject]$HT
    }

    $Document = $Word.Documents.Add()
    $Document.PageSetup.Orientation = 1
    $Document.Parent.Selection.ParagraphFormat.Alignment = 1
    $Picture = Get-ChildItem -Path $PictureFolder -Filter *.jpg | Get-Random
    $WordPic = $Document.InlineShapes.AddPicture($Picture.FullName)
    $Paragraph = $Document.Paragraphs.Add()
    $TableRange = @($Document.Paragraphs)[-1].Range
    $Headers = $TableContents[0].psobject.Properties.name
    $Columns = $Headers.count
    $Rows = $TableContents.shift.count + 1
    $Document.Parent.Selection.ParagraphFormat.Alignment = 1
    $Table = $Document.Tables.Add($TableRange,$Rows,$Columns,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent)
    $Table.Style = "Grid Table 5 Dark - Accent 1"
    $Table.Style.Table.Alignment = 1
    $Column = 1
    $Row = 1
    #Build header
    $Headers.ForEach{
        Write-Verbose "Adding $($_)"
        $Table.cell($Row,$Column).range.text = $_
        $Column++
    }  
    $Column = 1    
    #Add Data
    For ($i=0; $i -lt $TableContents.Count; $i++) {
        $Headers.ForEach{
            $Table.cell(($i+2),$Column).range.text = $TableContents[$i].$_
            $Column++
        }
        $Column = 1 
    }
    FitToPage
    $MonthName = (Get-Culture).DateTimeFormat.GetMonthName($Month)
    $Document.SaveAs("$RootFolder\Word\$Month. StandIndeling $MonthName $((Get-Date).Year) - $Day.docx",12)
    $Document.SaveAs("$RootFolder\PDF\$Month. StandIndeling $MonthName $((Get-Date).Year) - $Day.pdf",17)
    $Document.Close()
}
$Word.Quit()

#endregion Create Schedule in Word and PDF

#region Send CalendarInvites

$Message = @"
Beste {0},

Jullie zijn samen ingedeeld voor de stand. Als je niet kunt op deze datum/tijd stuur me dan een privé bericht via WhatsApp, dan probeer ik je shift te ruilen met iemand anders. Je kunt ook deze uitnodiging afwijzen en daarin aangeven welke zaterdagen je wel beschikbaar bent.

Met vriendelijke groet,

Michaja van der Zouwen
+316 517 660 75
"@

foreach ($Shift in $ShiftSchedule.Where{$_.Start -gt (Get-Date)})
{
    If ($Shift.Volunteers)
    {
        $Parameters = @{
            Start = $Shift.Start
            End = $Shift.End
            Summary = "Stand: $($Shift.Volunteers[0]) & $($Shift.Volunteers[1])"
            Calendar = 'Stand Leyenburg'
            Description = $Message -f "$($Shift.Volunteers[0].GivenName) & $($Shift.Volunteers[1].GivenName)"
            Location = 'Haage Markt'
            Attendees = $Shift.Volunteers.EmailAddress
        }
        If (Get-GoogleCalendarEvent -Calendar $Parameters.Calendar -Date $Parameters.Start -EndDate $Parameters.End)
        {
            Write-Warning "Event already exists: $($Parameters.Start) - $($Parameters.End.ToShortTimeString()). Skipping..."
        }
        elseif ($Parameters.Attendees.Count -lt 2)
        {
            Write-Error "Not enough volunteers: $($Parameters.Start) - $($Parameters.End.ToShortTimeString()). Skipping..."
        }
        else
        {
            New-GoogleCalendarEvent @Parameters -SendNotifications
            sleep -s 1
        }
    }
    else
    {
        Write-Warning "No Volunteers: $($Parameters.Start) - $($Parameters.End.ToShortTimeString()). Skipping..."
    }
}

#endregion Send CalendarInvites
