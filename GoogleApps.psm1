#region helper functions

# Get or refresh access token
function Get-GoogleAccessToken
{
    [CmdletBinding()]
    Param(
        # Google App to gain access to
        [Parameter(Mandatory=$true)]
        [GoogleApp]
        $App,

        # Refresh current token
        [switch]
        $Refresh
    )

    $ClientIDInfo = Get-ItemProperty -Path HKCU:\Software\GooglePoSH
    $TokenParams = @{
        client_id = $ClientIDInfo.client_id
        client_secret = $ClientIDInfo.client_secret
        grant_type = 'refresh_token';
    }
    If ($ClientIDInfo."$App`Token" -or $Refresh)
    {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR(($ClientIDInfo."$App`Token" | ConvertTo-SecureString))
        $TokenParams.Add('refresh_token',[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR))
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    }
    else
    {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR(($ClientIDInfo."$App`Code" | ConvertTo-SecureString))
        $TokenParams.Add('code',[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR))
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        $TokenParams.grant_type = 'authorization_code'
        $TokenParams.Add('redirect_uri',$ClientIDInfo.redirect_uris[0])
        Try
        {
            $Token = Invoke-WebRequest -Uri $ClientIDInfo.token_uri -Method POST -Body $TokenParams -ErrorAction Stop | ConvertFrom-Json
        }
        catch
        {
            Throw 'Unable to get authorization code. Please reset.'
        }
        $SecureToken = $Token.refresh_token | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
        $null = New-ItemProperty -Path HKCU:\Software\GooglePoSH -Name "$App`Token" -Value $SecureToken
        $TokenParams.Add('refresh_token',$Token.refresh_token)
        $TokenParams.Remove('code')
        $TokenParams.Remove('redirect_uri')
        $TokenParams.grant_type = 'refresh_token'
    }
    $RefreshedToken = Invoke-WebRequest -Uri "https://accounts.google.com/o/oauth2/token" -Method POST -Body $TokenParams | ConvertFrom-Json
    New-Variable -Name "$App`Access" -Scope global -Value @{
        access_token = $RefreshedToken.access_token
        expires = (Get-Date).AddSeconds($RefreshedToken.expires_in)
    } -Force -PassThru | select -ExpandProperty Value
}

# Invoke a Google API request
function Invoke-GoogleAPI
{
    [CmdletBinding()]
    Param(
        # Google App to connect to
        [Parameter(Mandatory=$true)]
        [GoogleApp]
        $App,

        [Microsoft.PowerShell.Commands.WebRequestMethod]
        $Method = 'Default',

        [Parameter(Mandatory=$true)]
        [string]
        $Target,

        [hashtable]
        $Options,

        $Body
    )

    $ClientIDInfo = Get-ItemProperty -Path HKCU:\Software\GooglePoSH
    $AccessToken = Get-Variable -Name "$App`Access" -ValueOnly -ErrorAction SilentlyContinue
    If (!$AccessToken -or ($AccessToken.expires -le (Get-Date)))
    {
        Write-Verbose 'Getting new access token'
        $AccessToken = Get-GoogleAccessToken -App $App -Refresh -ErrorAction Stop
        sleep -Milliseconds 500
    }

    $WRProperties = @{
        URI = "$($App.BaseURI)/$Target`?access_token=$($AccessToken.access_token)"
        Method = $Method
    }
    switch ($PsBoundParameters.Keys)
    {
        Body    {$WRProperties.Add('Body',$Body)}
        Options {
            $Options.GetEnumerator().ForEach{
                If ($_.Value -gt 1)
                {
                    foreach ($Value in $_.Value)
                    {
                        $WRProperties.URI = $WRProperties.URI + "&$($_.Key)=$Value"
                    }
                }
                else
                {
                    $WRProperties.URI = $WRProperties.URI + "&$($_.Key)=$($_.Value)"
                }
            }
        }
    }
    
    Invoke-WebRequest @WRProperties -ContentType 'application/json' | 
        Select -ExpandProperty Content | ConvertFrom-Json
}

# Set up a connection to the API of a Google app
function Connect-GoogleApp
{
    [CmdletBinding()]
    Param(
        # Google App to connect to
        [Parameter(Mandatory=$true)]
        [ValidateSet('GMail','Calendar','Contacts')]
        [GoogleApp]
        $App,

        # Path to the json file with the client secret
        [Parameter(Mandatory=$false)]
        $File = "O:\Mijn Documenten\Prive\Google_client_secret.json",

        # Reset and re-authorize the connection
        [switch]
        $Reset
    )
    $CodeName = "$App`Code"
    If ($ClientIDInfo = Get-ItemProperty -Path HKCU:\Software\GooglePoSH -ErrorAction SilentlyContinue)
    {
        If ($ClientIDInfo.$CodeName)
        {
            Write-Verbose "Retreiving authorization code [$CodeName] from registry..."
            $SecureCode = $ClientIDInfo.$CodeName
        }
    }
    elseif ($PSBoundParameters.File)
    {
        Write-Verbose "Creating new Authorization code using file [$File]"
        $ClientIDInfo = Get-Content $File -ErrorAction Stop | ConvertFrom-Json | select -ExpandProperty Installed
        
        $null = New-Item HKCU:\Software\GooglePoSH
    }
    else
    {
        Write-Error "No cached authorization found. Please provide a file containing a client secret." -Category NotSpecified
        return
    }
    If (!$SecureCode -or $Reset)
    {
        $Scope = $App.Scope
        $URL = "$($ClientIDInfo.auth_uri)?client_id=$($ClientIDInfo.client_id)" +`
                "&redirect_uri=$($ClientIDInfo.redirect_uris[0])" +`
                "&scope=$Scope&response_type=code"
        Start-Process $URL
        $SecureCode = Read-Host "Please enter your authorization code" -AsSecureString | 
            ConvertFrom-SecureString
        $null = New-ItemProperty -Path HKCU:\Software\GooglePoSH -Name $CodeName -Value $SecureCode -Force
        If ($ClientIDInfo."$App`Token")
        {
            Remove-ItemProperty -Path HKCU:\Software\GooglePoSH -Name "$App`Token"
        }
    }
    Get-GoogleAccessToken -App $App
}

#endregion helper functions

# Load classes
. "$PSScriptRoot\GoogleAppsClasses.ps1"

# Load calendar cmdlets
. "$PSScriptRoot\GoogleCalendar.ps1"
