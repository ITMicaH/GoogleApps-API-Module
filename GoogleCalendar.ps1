
function New-GoogleCalendarEvent
{
    [CmdletBinding()]
    Param(
        # Start date for the event
        [Parameter(Mandatory=$true)]
        $Start = '21-6-2018 12:00',

        # End date for the event
        $End = '21-6-2018 12:30',

        # Title for the event
        [Parameter(Mandatory=$true)]
        [string]
        $Summary,

        # Description for the event
        [string]
        $Description,

        # Location for the event
        [string]
        $Location,

        # Add attendees for the event
        [EventAttendee[]]
        $Attendees,

        # Reminder(s) for the event
        [EventReminder[]]
        $Reminder,

        # Event is all day
        [switch]
        $AllDay,

        # Send invite to attendees
        [switch]
        $SendNotifications
    )
    DynamicParam 
    {
        # Set the dynamic parameters' name
        $ParameterName = 'Calendar'

        # Create and set the parameters' attributes
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $false
        $ParameterAttribute.Position = 0

        # Generate and set the ValidateSet
        $arrSet = New-Object System.Collections.ArrayList
        $null = $arrSet.Add('Primary')
        $Calendars = @(Invoke-GoogleAPI -App Calendar -Method Get -Target users/me/calendarList).items
        $null = $arrSet.AddRange($Calendars.where{!$_.Primary}.Summary) #.foreach({$arrSet.Add($_.Summary)})
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)

        # Add the attributes to the attributes collection
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $AttributeCollection.Add($ParameterAttribute)
        $AttributeCollection.Add($ValidateSetAttribute)

        # Create and return the dynamic parameter
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
        return $RuntimeParameterDictionary
    }

    Process
    {
        If ($PSBoundParameters.Calendar)
        {
            $CalendarId = $Calendars.where{$_.Summary -eq $PSBoundParameters.Calendar}.Id
        }
        else
        {
            $CalendarId = 'primary'
        }
        If ($PSBoundParameters.End)
        {
            If ((Get-Date $Start -ErrorAction Stop) -gt (Get-Date $End -ErrorAction Stop))
            {
                Write-Error "Start date is later than end date." -Category InvalidArgument
            }
        }
        else
        {
            Write-Verbose 'No end date provided. Setting default end date.'
            $End = (Get-Date $Start).AddMinutes(30)
            $PSBoundParameters.end = $End
        }
        switch ($PSBoundParameters.AllDay)
        {
            $true   {$DateType = 'date';$DateFormat = 'yyyy-MM-dd'}
            Default {$DateType = 'dateTime';$DateFormat = 'yyyy-MM-ddTHH:mm:sszzz'}
        }
        Write-Verbose 'Creating json body'
        $ApiParam = @{
            App = 'Calendar'
            Method = 'POST'
            Target = "calendars/$CalendarId/events"
            Body = @{}
        }
        switch ($PSBoundParameters.keys)
        {
            Start {
                $ApiParam.Body.start = @{
                    $DateType = (Get-Date $Start -Format $DateFormat)
                }
            }
            End {
                $ApiParam.Body.end = @{
                    $DateType = (Get-Date $End -Format $DateFormat)
                }
            }
            Description {
                $ApiParam.Body.description = $Description
            }
            Location {
                $ApiParam.Body.location = $Location
            }
            Attendees {
                $ApiParam.Body.attendees = New-Object System.Collections.ArrayList
                $null = $Attendees.ForEach({$ApiParam.Body.attendees.Add($_)})
            }
            Summary {
                $ApiParam.Body.summary = $Summary
            }
            SendNotifications {
                $ApiParam.Add('Options',@{sendUpdates='all'})
            }
        }
        $ApiParam.Body = ($ApiParam.Body | ConvertTo-Json).Replace('null','""')
        
        Write-Verbose 'Invoking Google API'
        Invoke-GoogleAPI @ApiParam
    }
}


function Get-GoogleCalendarEvent
{
    [CmdletBinding()]
    Param(
        # (Start)Date for the event
        [Parameter()]
        $Date = '13-7-2019',

        # End date for the event
        $EndDate,

        # Search events for this string
        [string]
        $SearchString,

        # Reminder(s) for the event
        [hashtable]
        $Filter
    )
    DynamicParam 
    {
        # Set the dynamic parameters' name
        $ParameterName = 'Calendar'

        # Create and set the parameters' attributes
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $false
        $ParameterAttribute.Position = 0

        # Generate and set the ValidateSet
        $arrSet = New-Object System.Collections.ArrayList
        $null = $arrSet.Add('Primary')
        $Calendars = @(Invoke-GoogleAPI -App Calendar -Method Get -Target users/me/calendarList).items
        $null = $arrSet.AddRange($Calendars.where{!$_.Primary}.Summary) #.foreach({$arrSet.Add($_.Summary)})
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)

        # Add the attributes to the attributes collection
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $AttributeCollection.Add($ParameterAttribute)
        $AttributeCollection.Add($ValidateSetAttribute)

        # Create and return the dynamic parameter
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
        return $RuntimeParameterDictionary
    }

    Process
    {
        If ($PSBoundParameters.Calendar)
        {
            $CalendarId = $Calendars.where{$_.Summary -eq $PSBoundParameters.Calendar}.Id
        }
        else
        {
            $CalendarId = 'primary'
        }
        If ($PSBoundParameters.EndDate)
        {
            If ((Get-Date $Date -ErrorAction Stop) -gt (Get-Date $EndDate -ErrorAction Stop))
            {
                Write-Error "Start date is later than end date." -Category InvalidArgument -ErrorAction Stop
            }
        }
        elseif ($PSBoundParameters.Date)
        {
            Write-Verbose 'No end date provided. Setting default end date.'
            $End = (Get-Date $Date).AddDays(1)
            $PSBoundParameters.EndDate = $End
        }
        Write-Verbose 'Creating query url'
        $ApiParam = @{
            App = 'Calendar'
            Method = 'Get'
            Target = "calendars/$CalendarId/events/"
            Options = @{}
        }
        switch ($PSBoundParameters.keys)
        {
            Date {
                $ApiParam.Options.Add('timeMin',(Get-Date $Date -Format 'yyyy-MM-ddTHH:mm:ssZ'))
            }
            EndDate {
                $ApiParam.Options.Add('timeMax',(Get-Date $Date -Format 'yyyy-MM-ddTHH:mm:ssZ'))
            }
            SearchString {
                $ApiParam.Options.Add('q',$SearchString)
            }
            Filter {
                $ApiParam.Options.Add('sharedExtendedProperty',$Filter.GetEnumerator().ForEach{"$($_.Key)=$($_.Value)"})
            }
        }
        
        Write-Verbose 'Invoking Google API'
        Invoke-GoogleAPI @ApiParam | Select -ExpandProperty Items
    }
}

<#
    New-GoogleCalendarEvent -Start 22-6-2018 -Summary "Titel - TEST" -AllDay -Description "Test - description" -Location Testlocatie -Attendees m.vdzouwen@tweedekamer.nl -SendNotifications
#>
