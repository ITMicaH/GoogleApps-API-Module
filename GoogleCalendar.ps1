
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
        [Alias('Title')]
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
        [EventReminder]
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
        $ParameterName = 'CalendarId'

        # Create and set the parameters' attributes
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $false
        $ParameterAttribute.Position = 0

        # Generate and set the ValidateSet
        $arrSet = New-Object System.Collections.ArrayList
        $null = $arrSet.Add('Primary')
        $null = (Invoke-GoogleAPI -App Calendar -Method Get -Target users/me/calendarList).items.foreach({$arrSet.Add($_.Id)})
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
        If ($PSBoundParameters.CalendarId)
        {
            $CalendarId = $PSBoundParameters.CalendarId
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
            $End = (Get-Date $Start).AddMinutes(30)
            $PSBoundParameters.end = $End
        }
        switch ($PSBoundParameters.AllDay)
        {
            $true   {$DateType = 'date';$DateFormat = 'yyyy-MM-dd'}
            Default {$DateType = 'dateTime';$DateFormat = 'yyyy-MM-ddTHH:mm:sszzz'}
        }
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
                $ApiParam.Add('Options',"sendNotifications=$true")
            }
        }
        $ApiParam.Body = ($ApiParam.Body | ConvertTo-Json).Replace('null','""')

        Invoke-GoogleAPI @ApiParam
    }
}
