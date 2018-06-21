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
                $this.BaseURI = 'https://www.googleapis.com/gmail/v3/users/me'
            }
            Calendar {
                $this.Scope = 'https://www.googleapis.com/auth/calendar'
                $this.BaseURI = 'https://www.googleapis.com/calendar/v3/users/me'
            }
            Contacts {
                $this.Scope = 'https://www.google.com/m8/feeds/'
                $this.BaseURI = 'https://www.google.com/m8/feeds/'
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

enum APIMethod
{
    Default
    Get
    Head
    Post
    Put
    Delete
    Trace
    Options
    Merge
    Patch
}

#endregion Enums
