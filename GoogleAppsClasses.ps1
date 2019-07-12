#region Classes

class GoogleApp
{
    [GoogleAppName]$Name
    [uri]$Scope
    [uri]$BaseURI

    GoogleApp ([string]$Name)
    {
        $this.Name = $Name
        switch ($Name)
        {
            Gmail {
                $this.Scope = 'https://mail.google.com/'
                $this.BaseURI = 'https://www.googleapis.com/gmail/v3'
            }
            Calendar {
                $this.Scope = 'https://www.googleapis.com/auth/calendar'
                $this.BaseURI = 'https://www.googleapis.com/calendar/v3'
            }
            Contacts {
                $this.Scope = 'https://www.google.com/m8/feeds/'
                $this.BaseURI = 'https://www.google.com/m8/feeds'
            }
        }
    }
    [string] ToString ()
    {
        return $this.Name
    }
}

#region Calendar classes

# Class for Calendar event reminders
class EventReminder
{
   # Property with validate set
   [ValidateSet("email","sms","popup")]
   [string] $Method
   [int]    $Minutes

   # Constructor
   EventReminder ([string]$Method, [int]$Minutes)
   {
       $this.Method = $Method  
       $this.Minutes = $Minutes     
   }

   EventReminder ([hashtable]$Reminder)
   {
       $this.Method = $Reminder.Method  
       $this.Minutes = $Reminder.Minutes     
   }

   [string] ToString()
   {
       return "$($this.Method)($($this.Minutes))"
   }
}

# Class for Calendar event attendees
class EventAttendee
{
    [ValidatePattern('.+@.+\.\w+')]
    [string] $email
    [string] $displayName
    [int]    $additionalGuests = 0
    [bool]   $optional = $false
    [bool]   $resource = $false

    # Constructor email
    EventAttendee ([string]$Email)
    {
        $this.email = $Email
    }

    # Constructor email and displayname
    EventAttendee ([string]$Email,[string]$DisplayName)
    {
        $this.email = $Email
        $this.displayName = $DisplayName
    }

    # Constructor email and optional
    EventAttendee ([string]$Email,[bool]$Optional)
    {
        $this.email = $Email
        $this.optional = $Optional
    }

    # Constructor all
    EventAttendee ([string]$Email,[string]$DisplayName,
                   [int]$AdditionalGuests,[bool]$Optional,
                   [bool]$Resource)
    {
        $this.email = $Email
        $this.displayName = $DisplayName
        $this.additionalGuests = $AdditionalGuests
        $this.optional = $Optional
        $this.resource = $Resource
    }

    [string] ToString()
    {
        return $this.email
    }
}

#endregion Calendar classes

#endregion Classes

#region Enums

enum GoogleAppName
{
    GMail
    Calendar
    Contacts
}


#endregion Enums
