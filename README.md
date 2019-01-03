# PowerShell
A collection of various PowerShell scripts

## Update-MailChimp.ps1

A script to read data from a SQL Database and synchronize it with a MailChimp list.
It will add new members to the list, and update existing members (but won't change their subscription status).
The script includes detailed help text and supports -Verbose and -Debug parameters.

This script currently only reads the SQL columns Email, FirstName, LastName, Title, and AgreedToPromotions.
If you need to add additional columns, the process is:

1. *Optional* Add a parameter for the column name to the end of the `[CmdletBinding()]Param(` section.
```    
#Name of the column in the SQL table that contains the customer's Example. Default value is 'Example'.
    [Parameter(HelpMessage="Name of the column in the SQL table that contains the customer's Example.")]
    [ValidateScript({
        If ($_ -match '^[a-zA-Z0-9 \-\._#@]+$') {
            $True
        }
        else {
            Throw "$_ is not a valid name for a SQL Schema, Table or Column."
        }
        })]
    [string]$SQLColExample = "Example"  
```


2. Add your new column name parameter to the SQL query.
```
# Prepare the SQL Query
$SQLQuery = "SELECT [$SQLColEmail],
    [$SQLColFName],
    [$SQLColLName],
    [$SQLColTitle],
    [$SQLColAgreed],
    [$SQLColExample]
    FROM [$SQLSchema].[$SQLTable]
"
```   
Note that the last column name should not include a trailing comma.
If you did not create a parameter in step 1, just enter the column name in squre brackets. e.g. `[Example]`


3. Add the new column to $BodyMergeFields. The key name is the name of the column in the MailChimp list.
```    
# Build the API request body for this customer - fill out an array to convert to JSON later
    $BodyMergeFields = @{
        FNAME 	    = 	$CustData[$i].$SQLColFName;
        LNAME 	    = 	$CustData[$i].$SQLColLName;
        TITLE	    =	$CustData[$i].$SQLColTitle;
        AGREED	    =	$CustData[$i].$SQLColAgreed;
        EXAMPLE	    =	$CustData[$i].$SQLColExample;
    }
```
Again, if you did not create a parameter in step 1, just add the column name. e.g. `EXAMPLE	    =	$CustData[$i].Example`


4. If your new column is a date field you will need to format it correctly for MailChimp's API.
The following code can be included after $BodyMergeFields is defined:
```
    # Format the dates correctly. Check the date fields are not DBNull objects as this will cause an error. Set them to an empty string if so.
    If ( $BodyMergeFields.EXAMPLE.getType().Name -eq "DBNull" ) {
        $BodyMergeFields.EXAMPLE = ""
    } Else {
        $BodyMergeFields.EXAMPLE = $BodyMergeFields.DATE|Get-Date -UFormat %Y/%m/%d
    }
    
```
    
