#initialize the Outlook application
[Microsoft.Office.Interop.Outlook.Application] $outlook = New-Object -ComObject Outlook.Application

#store all values from the default GAL to $entries
$entries = $outlook.Session.GetGlobalAddressList().AddressEntries

#declare array outside of the loop 
$object = @()

#loop through all entries retrieved
foreach ($entry in $entries) {

#Set values from each object
$firstname = $entry.getExchangeUser().FirstName
$lastname = $entry.getExchangeUser().LastName
$email = $entry.getExchangeUser().PrimarySMTPAddress
$mobile = $entry.getExchangeUser().MobileNumber
$officel = $entry.getExchangeUser().OfficeLocation
$jobtitle = $entry.getExchangeUser().JobTitle
$officephone = $entry.getExchangeUser().OfficePhone
$city = $entry.getExchangeUser().City
$state =  $entry.getExchangeUser().stateorprovince	
$postalcode = $entry.getExchangeUser().PostalCode
$streetaddress = $entry.getExchangeUser().StreetAddress
$company = $entry.getExchangeUser().Company
$country = $entry.PropertyAccessor.GetProperty(‘http://schemas.microsoft.com/mapi/proptag/0x3A26001E’)

#Export a for loop to a CSV file
$object += New-Object -TypeName PSObject -Property @{
               FirstName = $firstname
               LastName = $lastname
                       Email = $email
                       Mobile = $mobvile
                       OfficeLocation = $officel
                       JobTitle = $jobtitle
                       OfficePhone = $officephone
                       City = $city
                       State = $state
                       PostalCode = $postalcode
                       StreetAddress = $streetaddress
                       Company = $company
                       Country = $Country
} | Select-Object FirstName, LastName, Email, Mobile, OfficeLocation, JobTitle, OfficePhone, City, State, PostalCode, StreetAddress, Company, Country

}

#exports object to file or an array to a file
$object | export-csv c:\temp\outlookaddresslist.csv -notypeinformation

You need to create a Temp folder on your C: drive for this to work.



Alternatively, if you have access to Exchange, this can be run:



$entries = (Get-GlobalAddressList 'Default Global Address List’).RecipientFilter

Get-Recipient -RecipientPreviewFilter $filter | Where-Object {$_.HiddenFromAddressListsEnabled -ne $false} | Select-Object Name,PrimarySmtpAddress, Mobile, OfficeLocation, JobTitle, OfficePhone, City, State, PostalCode, StreetAddress, Company, Country | Export-CSV c:\Temp\ExchangeAddressList.csv -NoTypeInformation

