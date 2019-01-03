<#  
.SYNOPSIS  
    Extracts customer data from a SQL database and uploads data to MailChimp.

.DESCRIPTION  
    This script will synchronize new customer data from a SQL database to MailChimp.
    It uses MailChimp's REST API v3 and SqlPS PowerShell module.
    API Documentation available at https://developer.mailchimp.com/documentation/mailchimp/guides/get-started-with-mailchimp-api-3/
    
    You must provide a MailChimp API Key and List ID, these can be retrieved from the Account or Lists section on MailChimp's website.
    You must also provide the SQL server, database and table/view name to extract data from.
    You may optionally provide a SQL instance name and a SQL schema that owns the table/view. If not specified, the default instance and [dbo] schema will be used.
    You may also specify column names to read from the SQL table/view. If not specified, default column names will be used.
    To view the default column names, see the detailed help for this cmdlet.

    The PUT method is used for the MailChimp API request. This will create new customers and update data of existing customers.
    Subscription status of existing customers won't be changed as this is better managed within Mailchimp.

    This script requires the following columns in the source database table/view:
    Email, FirstName, LastName, Title, AgreedToPromotions
    The script can be modified to include additional tables if required.

.NOTES
    Author :        Chris Byrne
    Version:        1.0
    Creation Date:  2019-01-03

.EXAMPLE
    Update-Mailchip -MCAPIKey xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx-xxx -MCListID xxxxxxxxxx -SQLServer sql01 -SQLDB CustomersDB -SQLTable CustomerData

    Basic example showing minimum required parameters to load data from dbo.CustomerData table in the CustomersDB database, and synchronize it with MailChimp.
    Windows Authentication is used to connect to the SQL server.

.INPUTS
    This script does not accept pipeline input.

.OUTPUTS
    PSObject - a count of records that are successfully updated, records that failed to update due to an error in the API response,
    and records that were skipped because they lacked an email address.

#>

[CmdletBinding()]
Param(
    #The MailChimp API Key - retrieve this from the Mailchimp website, under Account > Extras.
    [Parameter(Position=0, Mandatory=$true, HelpMessage="The MailChimp API Key - retrieve this from the Mailchimp website, under Account > Extras")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9\-\._]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid MailChimp API Key."
        }
        })]
    [string]$MCAPIKey,

    #The MailChimp API Base URL, including the version number. The <dc> component will be extracted from the API key. Defaults to "api.mailchimp.com/3.0".
    [Parameter(HelpMessage="The MailChimp API Base URL, including the version number. The <dc> component will be extracted from the API key.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9\-\./_]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid MailChimp API Base URL."
        }
        })]
    [string]$MCAPIURL = "api.mailchimp.com/3.0",

    #The MailChimp API Username. This is not checked by the API service at present, but can be changed here if necessary in the future. Defaults to "MCUser".
    [Parameter(HelpMessage="The MailChimp API Username. This is not checked by the API service at present, but can be changed here if necessary in the future.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9\-\./_]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid MailChimp user name."
        }
        })]
    [string]$MCUser = "MCUser",

    #The MailChimp List ID to update. Get this from Mailchimp, under Lists > 'List name' > Settings > List name and campaign defaults.
    [Parameter(Position=1, Mandatory=$true, HelpMessage="The MailChimp List ID to update. Get this from Mailchimp, under Lists > 'List name' > Settings > List name and campaign defaults.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9\-\._]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid MailChimp List ID."
        }
        })]
    [string]$MCListID,

    #The status used for new customers loaded to MailChimp. By default, new customers will be subscribed to the list. Existing customers won't be changed.
    [Parameter(HelpMessage="The status used for new customers loaded to MailChimp. By default, new customers will be subscribed to the list. Existing customers won't be changed.")]
    [ValidateSet('subscribed','unsubscribed','pending','cleaned')]
    [string]$MCCustStatus = "subscribed",
     
    #The hostname of the SQL Server to connect to for source data.
    [Parameter(Position=2, Mandatory=$true, HelpMessage="The hostname of the SQL Server to connect to for source data.")]
    [string]$SQLServer,

    #The name of the SQL Server instance, if not specified the default instance will be used.
    [Parameter(HelpMessage="The name of the SQL Server instance, if not specified the default instance will be used.")]
    [string]$SQLInstance,
    
    #The name of the SQL Database where the source data resides.
    [Parameter(Position=3, Mandatory=$true, HelpMessage="The name of the SQL Database where the source data resides.")]
    [string]$SQLDB,

    #Username for the SQL Server connection. You must also specify a password. If neither are specified, Windows authentication will be used.
    [Parameter(HelpMessage="Username for the SQL Server connection. You must also specify a password. If not specified, Windows authentication will be used.")]
    [string]$SQLUserName,

    #Password for the SQL Server connection. You must also specify a username. If neither are specified, Windows authentication will be used.
    [Parameter(HelpMessage="Password for the SQL Server connection. You must also specify a username. If not specified, Windows authentication will be used.")]
    [string]$SQLPassword,

    #Name of SQL Schema that owns the Table or View containing source data. If not specified, the default schema [dbo] will be used.
    [Parameter(HelpMessage="The name of SQL Schema that owns the Table or View containing source data.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9 \-\._#@]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid name for a SQL Schema, Table or Column."
        }
        })]
    [string]$SQLSchema ="dbo",

    #Name of SQL Table or View containing source data, within the specified Database.
    [Parameter(Position=4, Mandatory=$true, HelpMessage="The name of the SQL Table or View containing source data, within the specified Database.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9 \-\._#@]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid name for a SQL Schema, Table or Column."
        }
        })]
    [string]$SQLTable,

    #Name of the column in the SQL table that contains the customer's email address. Default value is 'Email'.
    [Parameter(Position=5, HelpMessage="Name of the column in the SQL table that contains the customer's email address.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9 \-\._#@]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid name for a SQL Schema, Table or Column."
        }
        })]
    [string]$SQLColEmail = "Email",

    #Name of the column in the SQL table that contains the customer's Given Name. Default value is 'FirstName'.
    [Parameter(Position=6, HelpMessage="Name of the column in the SQL table that contains the customer's Given Name.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9 \-\._#@]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid name for a SQL Schema, Table or Column."
        }
        })]
    [string]$SQLColFName = "FirstName",

    #Name of the column in the SQL table that contains the customer's Last Name. Default value is 'LastName'.
    [Parameter(Position=7, HelpMessage="Name of the column in the SQL table that contains the customer's Last Name.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9 \-\._#@]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid name for a SQL Schema, Table or Column."
        }
        })]
    [string]$SQLColLName = "LastName",
 
    #Name of the column in the SQL table that contains the customer's Title. Default value is 'Title'.
    [Parameter(Position=8, HelpMessage="Name of the column in the SQL table that contains the customer's Title.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9 \-\._#@]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid name for a SQL Schema, Table or Column."
        }
        })]
    [string]$SQLColTitle = "Title",  

    #Name of the column in the SQL table that contains the customer's Agreed to Promotions response. Default value is 'AgreedToPromotions'.
    [Parameter(Position=9, HelpMessage="Name of the column in the SQL table that contains the customer's Agreed to Promotions response.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9 \-\._#@]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid name for a SQL Schema, Table or Column."
        }
        })]
    [string]$SQLColAgreed = "AgreedToPromotions"  

)

$ProgressPreference = 'SilentlyContinue' # Hides the progress bar for Invoke-WebRequest, gives a big performance boost

Write-Verbose "Loading the SQLPS module"
Try {
    Push-Location
    Import-Module SqlPS -DisableNameChecking  -ErrorAction Stop
    Pop-Location
}
Catch {
    Write-Verbose "An error occured loading the SQLPS Module. Please ensure you have it installed."
    Throw $Error[0]
}


# If an instance for SQL server was specified, concatenate the server and instance for invoke-sqlcmd. 
If ( $SQLInstance -ne "" -and $SQLInstance -ne $null ) {
    $SQLServerInstance = "$SQLServer\$SQLInstance"
} Else {
    $SQLServerInstance = $SQLServer
}


# Prepare the SQL Query
$SQLQuery = "SELECT [$SQLColEmail],
    [$SQLColFName],
    [$SQLColLName],
    [$SQLColTitle],
    [$SQLColAgreed]
    FROM [$SQLSchema].[$SQLTable]
"

Write-Verbose "Connecting to the SQL server $SQLServerInstance and pulling data from $SQLSchema.$SQLDB"
Write-Debug "SQL Query: $SQLQuery"

Try {
    If (( $SQLUserName -ne "" -and $SQLUserName -ne $null ) -and ( $SQLPassword -ne "" -and $SQLPassword -ne $null )) {
        Write-Verbose "Authenticating to SQL server as $SQLUserName"
        $CustData = @(Invoke-Sqlcmd -ServerInstance $SQLServerInstance -Database $SQLDB -Query $SQLQuery -Username $SQLUserName -Password $SQLPassword -WarningAction Ignore -ErrorAction Stop)
    } Else {
        Write-Verbose "Username and password were not specified, using Windows authentication."
        $CustData = @(Invoke-Sqlcmd -ServerInstance $SQLServerInstance -Database $SQLDB -Query $SQLQuery -WarningAction Ignore -ErrorAction Stop)
    }   
}
Catch {
    Write-Verbose "An error occured connecting to the SQL server. Please ensure it is online and not blocked by firewall, and you have adequate permissions to the database and table.
        Please also ensure you have specified the column names correctly."
    Throw $Error[0]
}

# Check that we actually retrieved some data
Write-Debug ([string]$CustData.Count + " Rows retrieved from the database" )
If ( $CustData.Count -lt 1 ) {
    Throw "A connection to the SQL Server was made but no data was retrieved. Please check permissions to the table, and that the table specified is correct."
}


# Build the authentication headers for the MailChimp API
$MCHeaders = @{ Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${MCUser}:${MCAPIKey}")) }
Write-Debug ("Authentication header: " + $MCHeaders.Authorization)

# Get the MailChimp API URL
$MCAPIURL = $MCAPIURL -replace "http[s]?://" # strip HTTP/s
$MCAPIURL = $MCAPIURL -replace "/$" # strip trailing slash

# Work out the <dc> component from the API key. This is delimited by a hyphen at the end of the key. Skip if the user specified this already.
If ( -Not ($MCAPIURL -imatch "^[a-z][a-z][0-9][0-9]?\." ) ) {
    $MCAPIURL = ($MCAPIKey.Split("-")[-1]) + "." + $MCAPIURL
}
Write-Debug "Using API URL: $MCAPIURL"


# Check the specified MailChimp list exists.
Try {
    $MCListName = Invoke-RestMethod -Uri ( "https://" + $MCAPIURL  + "/lists/" + $MCListID ) -Method Get -Headers $MCHeaders -ErrorAction Stop -WarningAction Ignore|Select -ExpandProperty Name
}
Catch {
    Write-Verbose "An error occured connecting to the MailChimp API and retrieving the list details.
        Please ensure you have specified the API key, API URL, and the List ID correctly. "
    Throw $Error[0]
}

Write-Debug "MailChimp list found: $MCListName"

# Set up the MD5 conversion objects
$md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
$utf8 = New-Object -TypeName System.Text.UTF8Encoding



$Results = [ordered]@{ Successful = 0; Failed = 0; Skipped = 0 } # Counts results of each API call for output at end of command.

# Loop through each data row in CustData, build the request body and call the API to update each customer.
$i = 0 
For ( $i; $i -lt $CustData.Length; $i++ ) {
    Write-Debug "Data row $i"
    # Check the customer has an email address - if not exit this loop and continue with the next record in the for loop
    If ( $CustData[$i].$SQLColEmail.GetType().Name -eq "DBNull" ) {
        Write-Verbose  ("Customer " + $CustData[$i].$SQLColFName,  $CustData[$i].$SQLColLName + " at database row $i skipped as they have no email address.")
        $Results.Skipped += 1
        Continue
    }

    # Build the MD5 hash of the email address, for the Mailchimp Customer ID
    $CustHash = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($CustData[$i].$SQLColEmail.toLower().trim()))).tolower() -replace "-"

    # Build the API request body for this customer - fill out an array to convert to JSON later
    $BodyMergeFields = @{
        FNAME 	    = 	$CustData[$i].$SQLColFName;
        LNAME 	    = 	$CustData[$i].$SQLColLName;
        TITLE	    =	$CustData[$i].$SQLColTitle;
        AGREED	    =	$CustData[$i].$SQLColAgreed;
    }
    # BodyMergeFields is a nested array within BodyData
    $BodyData = @{ 
        email_address   =   $CustData[$i].$SQLColEmail.Trim();
        status_if_new   =   $MCCustStatus.ToLower();
        merge_fields    =   $BodyMergeFields 
    }

    # Format the dates correctly. Check the date fields are not DBNull objects as this will cause an error. Set them to an empty string if so.
    # This is currently unused but is left here in case we need to sync any date fields in the future.
    #If ( $BodyMergeFields.DATE.getType().Name -eq "DBNull" ) {
    #    $BodyMergeFields.DATE = ""
    #} Else {
    #    $BodyMergeFields.DATE = $BodyMergeFields.DATE|Get-Date -UFormat %Y/%m/%d
    #}
    

    
    # Finally, call the API to update this customer
    Write-Verbose ( "Updating "+  $BodyData.email_address + " - Sending update request to https://" + $MCAPIURL  + "/lists/" + $MCListID + "/members/" + $CustHash )
    Write-Debug ( "Request body: " + ($BodyData|ConvertTo-JSON) )
    
    $RequestStatus = Try  {
        (Invoke-WebRequest -Uri ( "https://" + $MCAPIURL  + "/lists/" + $MCListID + "/members/" + $CustHash ) -Method PUT -Headers $MCHeaders -Body ($BodyData|ConvertTo-JSON) -ErrorAction Stop).BaseResponse
    } Catch [System.Net.WebException] {
        Write-Verbose "An exception was caught while performing the API request: $($_.Exception.Message)"
        Write-Debug $_.ErrorDetails.Message
        $Results.Failed += 1
    }
    If  ($RequestStatus.StatusCode.Value__ -eq "200" ) { 
        $Results.Successful += 1
        Write-Debug ("Update successful. Status: " + $RequestStatus.StatusCode.Value__ + " " +  $RequestStatus.StatusCode )
    }
}

# Update has completed, output our count of results and exit
[PSCustomObject]$Results|Write-Output