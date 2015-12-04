#################################################################
#Created: Austin Heyne											#
#Date: June 2012												#
#Contact: aheyne@ou.edu											#
#																#
#Info: For help use get-help.									#
#																#
#################################################################

<#
	
.SYNOPSIS
	Script used to create alumni.ou.edu mailboxes from the gratuate spreadsheet.
.DESCRIPTION
	Options
	
	Run Full Automated
		Will run all processes required to complete an import of data from an excel spreadsheet.
		Interactivity is still required for aditional information.
	
	Update Excel Data
		This will import the necessary data from an excel spreadsheet and upload it to the apropriate SQLDB in alumnimail.
		The excel spreadsheet needs to be in the standard format provided. 
	
	Update AD Data
		This will load the active directory ad module so this needs to be present on the machine running the script.
		It will also download additional data about the graduated students provided for in the excel spreadsheet and upload this data to the correct SQLDB
	
	Update Live Data
		This will connect to a Microsoft Live remote posh session and dowload data on currently created alumni mailboxes and update the SQLDB. 
		
	Generate Missing Live Emails
		Used to create any pending mailboxes listed in the SQLDB. This can be used when manually running the script or if continuing after an error.
		
.NOTES
	There are not parameters for the script. If any information is needed your will be prompted for it.
		
#>

function excel-import ($excelpath, $excelpsw){
	Write-Host "Starting Excel Import..."
	#start up excel silently
	$excel = New-Object -com excel.application
	$excel.visible = $False
	#import new workbook and decrypt
	$workbook = $excel.workbooks.open("$excelpath",1,$true,5,$excelpsw)
	$excelpsw = $null
	
	Write-Host "Import Successful, processing data, this will take a while..." 
#Process Sheet 1
	$sheet1 = $workbook.worksheets.item(1)
	$sheet1.saveas("C:\sheet1.csv",6)
	$sheet1 = $null
	$sheet1csv = Import-Csv "C:\sheet1.csv"
	
	Write-Progress "Importing" "Complete:" -PercentComplete $(0*100)
	$global:exceldata.seniorids = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.seniorids += [int]$_.ID}else{$global:exceldata.seniorids += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(1/21*100)
	$global:exceldata.seniorfullname = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.seniorfullname += [int]$_.FULL_NAME_LFMI}else{$global:exceldata.seniorfullname += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(2/21*100)
	$global:exceldata.seniorfirstname = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.seniorfirstname += [int]$_.FIRST_NAME}else{$global:exceldata.seniorfirstname += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(3/21*100)
	$global:exceldata.seniormiddlename = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.seniormiddlename += [int]$_.MIDDLE_NAME}else{$global:exceldata.seniormiddlename += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(4/21*100)
	$global:exceldata.seniorlastname = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.seniorlastname += [int]$_.LAST_NAME}else{$global:exceldata.seniorlastname += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(5/21*100)
	$global:exceldata.senioracademicperiod = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.senioracademicperiod += [int]$_.ACADEMIC_PERIOD}else{$global:exceldata.senioracademicperiod += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(6/21*100)
	$global:exceldata.senioracademicperioddesc = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.senioracademicperioddesc += [int]$_.ACADEMIC_PERIOD_DESC}else{$global:exceldata.senioracademicperioddesc += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(7/21*100)
	$global:exceldata.senioremailou = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.senioremailou += [int]$_.EMAIL_OU}else{$global:exceldata.senioremailou += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(8/21*100)
	$global:exceldata.senioremailpreferredcode = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.senioremailpreferredcode += [int]$_.EMAIL_PREFERRED_CODE}else{$global:exceldata.senioremailpreferredcode += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(9/21*100)
	$global:exceldata.senioremailpreferredaddress = @()
	$sheet1csv | ForEach-Object {if($_ -ne $null){$global:exceldata.senioremailpreferredaddress += [int]$_.EMAIL_PREFERRED_ADDRESS}else{$global:exceldata.senioremailpreferredaddress += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(10/21*100)
	$sheet1csv = $null
	
	Write-Host "Working: Seniors Sheet done..."

#Process Sheet 2
	$sheet2 = $workbook.worksheets.item(2)
	$sheet2.saveas("C:\sheet2.csv",6)
	$sheet2 = $null
	$sheet2csv = Import-Csv "C:\sheet2.csv"

	Write-Progress "Importing" "Complete:" -PercentComplete $(11/21*100)
	$global:exceldata.gradids = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.gradids += [int]$_.ID}else{$global:exceldata.gradids += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(12/21*100)
	$global:exceldata.gradfullname = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.gradfullname += [int]$_.FULL_NAME_LFMI}else{$global:exceldata.gradfullname += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(13/21*100)
	$global:exceldata.gradfirstname = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.gradfirstname += [int]$_.FIRST_NAME}else{$global:exceldata.gradfirstname += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(14/21*100)
	$global:exceldata.gradmiddlename = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.gradmiddlename += [int]$_.MIDDLE_NAME}else{$global:exceldata.gradmiddlename += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(15/21*100)
	$global:exceldata.gradlastname = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.gradlastname += [int]$_.LAST_NAME}else{$global:exceldata.gradlastname += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(16/21*100)
	$global:exceldata.gradacademicperiod = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.gradacademicperiod += [int]$_.ACADEMIC_PERIOD}else{$global:exceldata.gradacademicperiod += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(17/21*100)
	$global:exceldata.gradacademicperioddesc = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.gradacademicperioddesc += [int]$_.ACADEMIC_PERIOD_DESC}else{$global:exceldata.gradacademicperioddesc += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(18/21*100)
	$global:exceldata.grademailou = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.grademailou += [int]$_.EMAIL_OU}else{$global:exceldata.grademailou += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(19/21*100)
	$global:exceldata.grademailpreferredcode = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.grademailpreferredcode += [int]$_.EMAIL_PREFERRED_CODE}else{$global:exceldata.grademailpreferredcode += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(20/21*100)
	$global:exceldata.grademailpreferredaddress = @()
	$sheet2csv | ForEach-Object {if($_ -ne $null){$global:exceldata.grademailpreferredaddress += [int]$_.EMAIL_PREFERRED_ADDRESS}else{$global:exceldata.grademailpreferredaddress += $null}}
	Write-Progress "Importing" "Complete:" -PercentComplete $(21/21*100)
	$sheet2csv = $null
	
	Write-Host "Working: Grad Sheet done; Cleaning up..."
	
	$excel.quit()
	del "C:\sheet1.csv"
	del "C:\sheet2.csv"
	
	if($global:exceldata.seniorids.lenght -eq $null -and $global:exceldata.gradids.length -eq $null){
		Write-Host -foregroundcolor Red -backgroundcolor Black "The Excel import seems to have failed."
	}
	
	return
}

function excel-upload (){
	Write-Host "Uploading Seniors to SQL DB this may take a while..."
	foreach ($i in $global:exceldata.seniorids){
		$Query = "insert into SeniorsOnly values ('$($global:exceldata.seniorids[$i - 1])', '$($global:exceldata.seniorfullname[$i - 1])', '$($global:exceldata.seniorfirstname[$i - 1])', '$($global:exceldata.seniormiddlename[$i - 1])', '$($global:exceldata.seniorlastname[$i - 1])', '$($global:exceldata.senioracademicperiod[$i - 1])', '$($global:exceldata.senioracademicperioddesc[$i - 1])', '$($global:exceldata.senioremailou[$i - 1])', '$($global:exceldata.senioremailpreferredcode[$i - 1])', '$($global:exceldata.senioremailpreferredaddress[$i - 1])')"
		$Timeout = '30'
		Query-Sql $Query $Timeout 
		Write-Progress "Uploading..." "Complete: $([Math]::Round($i/$t*100,2))%" -PercentComplete $($i / $t * 100)
	}
	Write-Host "Uploading Grad to SQL DB this may take a while..."
	foreach ($i in $global:exceldata.gradids){
		$Query = "insert into GradandLaw values ('$($global:exceldata.gradids[$i - 1])', '$($global:exceldata.gradfullname[$i - 1])', '$($global:exceldata.gradfirstname[$i - 1])', '$($global:exceldata.gradmiddlename[$i - 1])', '$($global:exceldata.gradlastname[$i - 1])', '$($global:exceldata.gradacademicperiod[$i - 1])', '$($global:exceldata.gradacademicperioddesc[$i - 1])', '$($global:exceldata.grademailou[$i - 1])', '$($global:exceldata.grademailpreferredcode[$i - 1])', '$($global:exceldata.grademailpreferredaddress[$i - 1])')"
		$Timeout = '30'
		Query-Sql $Query $Timeout
		Write-Progress "Uploading..." "Complete: $([Math]::Round($i/$t*100,2))%" -PercentComplete $($i / $t * 100)
	}	
}

function live-connect (){
	$Session = $null
	$import = $null
	Write-Host "Please supply Live connection credentials"
	$LiveCred = Get-Credential
	Write-Host "Attempting Connection..."
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
	if($Session -eq $null){
		Write-Host -foregroundcolor Red -backgroundcolor Black "Connecting to ps.outlook.com seems to have failed"
		main
	}
	Write-Host "Importing Powershell Session"
	$import = Import-PSSession $Session –AllowClobber
	if($import -eq $null){
		Write-Host -foregroundcolor Red -backgroundcolor Black "Powershell Session import seems to have failed"
		main
	}
	Write-Host "Live connection established."
	
	return $import
}

function live-download () {
	Write-Host "Initiating Live data download..."
	$global:livedata.san = @()
	$global:livedata.upn = @()
	$global:livedata.wli = @()
	$global:livedata.name = @()
	
	get-mailbox | ForEach-Object {
		Write-Progress "Downloading..." "Complete: $([Math]::Round(0/4*100,2))%" -PercentComplete $(0/4*100)
		if($_.samaccountname -ne $null){$global:livedata.san += $_.samaccountname}else{$global:livedata.san += $null}
		Write-Progress "Downloading..." "Complete: $([Math]::Round(1/4*100,2))%" -PercentComplete $(1/4*100)
		if($_.userprincipalname -ne $null){$global:livedata.upn += $_.userprincipalname}else{$global:livedata.upn += $null}
		Write-Progress "Downloading..." "Complete: $([Math]::Round(2/4*100,2))%" -PercentComplete $(2/4*100)
		if($_.windowsliveid -ne $null){$global:livedata.wli += $_.windowsliveid}else{$global:livedata.wli += $null}
		Write-Progress "Downloading..." "Complete: $([Math]::Round(3/4*100,2))%" -PercentComplete $(3/4*100)
		if($_.name -ne $null){$global:livedata.name += $_.name}else{$global:livedata.name += $null}
		Write-Progress "Downloading..." "Complete: $([Math]::Round(4/4*100,2))%" -PercentComplete $(4/4*100)
	}
	
	if($global:livedata.name.length -eq $null){
		Write-Host -foregroundcolor Red -backgroundcolor Black "The Download seems to have failed..."
		Main
	}
	
	Write-Host "Download complete."
	
	return
	
}

function live-upload () {
	Write-Host "Initiating Live data upload..."
	$i=1
	$t = $global:livedata.san.length
	Write-Host "Uploading data to SQLDB this may take a while..."
	while ($i -le $t){
		Write-Progress "Uploading..." "Complete: $([Math]::Round($i/$t*100,2))%" -PercentComplete $($i/$t*100)
		$Query = "insert into dbo.LiveData select '$($global:livedata.name[$i - 1])','$($global:livedata.san[$i - 1])','$($global:livedata.upn[$i - 1])','$($global:livedata.wli[$i - 1])' where not exists (select * from dbo.LiveData where SamAccountName='$($global:livedata.san[$i - 1])')"
		$Timeout = '30'
		Query-Sql $Query $Timeout
		$i += 1
	}
	Write-Host "Live data upload complete."
}

function live-createmailboxes () {
	increment-names
	$i = $l = 0
	$t = $global:createme.dotname.length
	while ($i -lt $t){
		$orgunit = 
		$san = 
		$upn = $global:createme.dotname[$i] + '@alumni.ou.edu'
		Write-Progress "Creating mailboxes and sending welcome messages..." "Complete: $([Math]::Round($i/$t*100,2))%" -PercentComplete $($i/$t*100)
		New-Mailbox -Name $global:createme.dotname[$i] -Alias $global:createme.dotname[$i] -OrganizationalUnit $orgunit -UserPrincipalName $upn -SamAccountName $san -FirstName $global:createme.firstname[$i] -Initials $global:createme.middlename[$i][0] -LastName $global:createme.lastname[$i] -Password $global:createme.fourbyfour[$i] -ResetPasswordOnNextLogon $true 
		get-mailbox $global:createme.dotname | set-mailbox -IssueWarningQuota '9 GB (9,663,676,416 bytes)' -ProhibitSendReceiveQuota '10 GB (10,737,418,240 bytes)'
		send-welcome $i
		$l+=1
	}
	Write-Host '$l mailboxes created.'
	return
}

function send-welcome ($i) {
	$Message = "Your Alumni email account is now ready for use. You may access your account by going to live.com. <br/>Username: " + $global:createme.dotname[$i] + "<br/>Password: " + $global:createme.fourbyfour[$i] + "<br/>Thank You"
	$SmtpClient = new-object system.net.mail.smtpClient 
	$SmtpClient.Host = "smtp.ou.edu"
	$address = $global:createme.email[$i]
	$mailmessage = New-Object system.net.mail.mailmessage 
	$mailmessage.from = "postmaster@ou.edu"
	$mailmessage.To.add($address)
	$mailmessage.Subject = "Welcome to Alumni Mail"
	$mailmessage.IsBodyHtml = $False
	$mailmessage.Body = $Message
	$mailmessage.Priority = "Normal"
	$smtpclient.Send($mailmessage) 
	return
}

function increment-names (){
	Write-Host "Filtering and checking for duplicate names..."
	$i = 0 
	$t = $global:createme.dotname.length
	#loop through all to be created
	while ($i -lt $t) {
		Write-Progress "Duplicate checking names..." "Complete: $([Math]::Round($i/$t*100,2))%" -PercentComplete $($i/$t*100)
		#loop through all existing
		$j = 0
		$k = $global:livedata.name.lenght
		while ($j -lt $k) {
			if ($global:createme.dotname[$i] -match $global:livedata.name[$j]) {
				Write-Host 'Duplicate found at $global:createme.dotname[$i], incrementing...'
				$c = $global:createme.dotname[$i].length
				$global:createme.dotname[$i] = $global:createme.dotname[$i].substring(0,$c-1) + $([System.Double]::Parse($global:createme.dotname[$i][$c-1]) + 1).tostring()
				$j=-1
			}
			$j+=1
		}
		$i+=1
	}
	return
}

function New-SqlConnection () {
	$SqlServer = "it-central.sooner.net.ou.edu"
	$SqlDBName = "Alumnimail"
	#Close Connection by namd $ConnectionName. Allows refresh of connection.
	if(Test-Path variable:\$global:SQLConnection){
		Write-Host "Sql connection already active..."
		return
	}
	Write-Host "Initiating SQLDB connection..."
	$global:SQLConnection = New-Object system.Data.SqlClient.SqlConnection
	$global:SQLConnection.ConnectionString = "Server=$SqlServer;Database=$SQLDBName;Trusted_Connection=True"
	$global:SQLConnection.open()
	if(Test-Path variable:\$global:SQLConnection){
		Write-Host "Sql connection established."
	}
	
	return
}

function Close-SqlConnection () {
	$global:SQLConnection.close()
	$global:SQLConnection = $null
}

function Close-SqlQuery ($Query) {
	$Query.close()
	$Query = $null
}

function Query-Sql ($Query, [int]$CommandTimeout) {
	#$Query is sqlcommand text; $sqlconnection is connection object, used to make command variable
	$cmd = $global:SQLConnection.CreateCommand()
	$cmd.commandtext = $Query
	$cmd.commandtimeout = $CommandTimeout
	$dataadapter = New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
	$DataSet = New-Object System.Data.DataSet
	$dataadapter.fill($DataSet)
	
	$cmd = $null
	return $DataSet
}

function ad-download () {
	Write-Host "Colecting AD Data..."
	$global:addata.ouid = @()
	$global:addata.name = @()
	$global:addata.fourbyfour = @()
	$all = @()
	$i = 0
	$all = get-aduser -filter * -properties employeenumber,extensionattribute6,samaccountname
	$t = $all.length
	while($i -lt $t){
		Write-Progress "Filtering and compiling data..." "Complete: $([Math]::Round($i/$t*100,2))%" -PercentComplete $($i/$t*100)
		if(($all[$i].samaccountname -ne $null) -and ($all[$i].samaccountname -match "[a-z][a-z][a-z0-9][a-z0-9][0-9][0-9]") -and ($all[$i].extensionattribute6 -ne $null)){
			$global:addata.fourbyfour += $all[$i].samaccountname
			$global:addata.name += $all[$i].extensionattribute6
			if($all[$i].employeenumber -ne $null){$global:addata.ouid += $all[$i].employeenumber}else{$global:addata.ouid += $null}
		}
		$i+=1
	}	
}

function ad-upload () {
	Write-Host "Initiating AD data upload..."
	$i=1
	$t = $global:addata.length
	Write-Host "Uploading data to SQLDB this may take a while..."
	while ($i -le $t){
		Write-Progress "Uploading..." "Complete: $([Math]::Round($i/$t*100,2))%" -PercentComplete $($i/$t*100)
		$global:addata.fourbyfour
		$Query = "insert into dbo.AdData select '$($global:addata.fourbyfour[$i - 1])','$($global:AdData.ouid[$i - 1])','$($global:AdData.name[$i - 1])'"
		$Timeout = '30' 
		Query-Sql $Query $Timeout
		$i += 1
	}
	Write-Host "AD data upload complete."
}

function sql-jobs () {
	Write-Host "Making temporary SQL table..."
	$Query = "create table #temp (ouid int, dotname varchar(MAX), fourbyfour varchar(8), firstname varchar(MAX), middlename varchar(MAX), lastname varchar(MAX), email varchar(MAX) )"
	$Timeout = '30' 
	Query-Sql $Query $Timeout
	Write-Host "Moving SQL Data..."
	$Query = "insert into #temp (ouid, firstname, middlename, lastname, email) select ID, FIRST_NAME, MIDDLE_NAME, LAST_NAME, EMAIL_PREFERRED_ADDRESS from dbo.SeniorsOnly"
	$Timeout = '30' 
	Query-Sql $Query $Timeout
	$Query = "insert into #temp (ouid, firstname, middlename, lastname, email) select ID, FIRST_NAME, MIDDLE_NAME, LAST_NAME, EMAIL_PREFERRED_ADDRESS from dbo.GradandLaw"
	$Timeout = '30' 
	Query-Sql $Query $Timeout
	$Query = "insert into #temp (dotname, fourbyfour) select dotname, fourbyfour from dbo.AdData where soonerid in(select ouid from bla)"
	$Timeout = '30' 
	Query-Sql $Query $Timeout
	Write-Host "SQL moving done."
}

function sql-download () {
	$global:createme.ouid = @()
	$global:createme.dotname = @()
	$global:createme.fourbyfour = @()
	$global:createme.firstname = @()
	$global:createme.middlename = @()
	$global:createme.lastname = @()
	$global:createme.email = @()
	
	$Query = "select * fron #temp"
	$Timeout = '30'
	Write-Host "Retrieving Sql data..."
	$temp = Query-Sql $Query $Timeout
	$i = 0
	$t = $temp.tables[0].rows.count
	while ($i -lt $t) {
		Write-Progress "Loading Data..." "Complete: $([Math]::Round($i/$t*100,2))%" -PercentComplete $($i / $t * 100)
		$global:createme.ouid[$i] = $temp.tables[0].rows[$i].ouid
		$global:createme.dotname[$i] = $temp.tables[0].rows[$i].dotname
		$global:createme.fourbyfour[$i] = $temp.tables[0].rows[$i].fourbyfour
		$global:createme.firstname[$i] = $temp.tables[0].rows[$i].firstname
		$global:createme.middlename[$i] = $temp.tables[0].rows[$i].middlename
		$global:createme.lastname[$i] = $temp.tables[0].rows[$i].lastname
		$global:createme.email[$i] = $temp.tables[0].rows[$i].email
	}
	Write-Host "SQL download complete."
}

#Initiation Functions

function MainPrompt (){
	$R = ([System.Management.Automation.Host.ChoiceDescription]"&Run Full Automated")
	$R.helpmessage = "Run all processes automatically."
	$E = ([System.Management.Automation.Host.ChoiceDescription]"&Excell Data Update")
	$E.helpmessage = "Imports and uploads an Excel Spreadsheet to SQLDB."
	$A = ([System.Management.Automation.Host.ChoiceDescription]"&AD Data Update")
	$A.helpmessage = "Downloads and uploads AD Data to SQLDB."
	$L = ([System.Management.Automation.Host.ChoiceDescription]"&Live Data Update")
	$L.helpmessage = "Downloads and uploads Live Data to SQLDB"
	$G = ([System.Management.Automation.Host.ChoiceDescription]"&Generate Missing Live Emails")
	$G.helpmessage = "Creates missing mailboxes from SQL Data. Use to update or correct creation error."
	$Q = ([System.Management.Automation.Host.ChoiceDescription]"&Quit")
	$Q.helpmessage = "Quit"

	$Caption = "Main Menu"
	$Message = "What function do you want to perform?"
	$Choices = ($R,$E,$A,$L,$G,$Q)
	$host.ui.PromptForChoice($Caption,$Message,[System.Management.Automation.Host.ChoiceDescription[]]$Choices,5)

}

function process-automated () {
	process-excel
	process-ad
	process-live
}

function process-excel () {
	$excelpath = Read-Host "What is the path to the excel file?"
	#note this plain text
	$excelpsw = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((read-host -assecurestring "What is the password to the excel file?")))
	excel-import $excelpath $excelpsw
	if(!$global:exceldata -eq $null){
		Write-Host "Excel data import appears to have completed successfully."
	}else{
		Write-Host -foregroundcolor Red -backgroundcolor Black "Something went wrong or the excel file is empty."
	}

	if(!$global:exceldata -eq $null){
		Write-Warning "You need to run Excel Import First"
	}else{
		new-sqlconnection
		if(!$global:SQLConnection -eq $null){
			Write-Host "Gathering error checking data..."
			$Query = "select count(ID) from SeniorsOnly"
			$Timeout = '30'
			$Seniorprelength = Query-Sql $Query $Timeout
			$Query = "select count(ID) from GradandLaw"
			$Gradprelength = Query-Sql $Query $Timeout
			Write-Host "Data gathered."
			
			excel-upload
			
			$Query = "select count(ID) from SeniorsOnly"
			$Timeout = '30'
			$Seniorpostlength = Query-Sql $Query $Timeout
			$Query = "select count(ID) from GradandLaw"
			$Gradpostlength = Query-Sql $Query $Timeout
			
			if($Seniorpostlength -ne $($Seniorprelength + $global:exceldata.seniorids.length)){
				$Errorspot = $Seniorpostlength - $Seniorprelength
				Write-Host -foregroundcolor Red -backgroundcolor Black "There appears to have been an error around line $Errorspot in the Senior excel file."
			}else{
				Write-Host "Senior excel file seems to have uploaded fine."
			}
			if($Gradpostlength -ne $($Gradprelength + $global:exceldata.gradids.length)){
				$Errorspot = $Gradpostlength - $Gradprelength
				Write-Host -foregroundcolor Red -backgroundcolor Black "There appears to have been an error around line $Errorspot in the Grad excel file."
			}else{
				Write-Host "Grad excel file seems to have uploaded fine."
			}
		}else{
			Write-Host -foregroundcolor Red -backgroundcolor Black "SQLDB Connection Failed."
		}
	}
	
	Main 
}

function process-ad () {
	Import-Module activedirectory
	ad-download
	ad-upload
}

function process-live () {
	live-connect
	live-download
	live-upload
}

function process-generateemails () {
	live-createmailboxes
}

function process-quit () {
	Write-Host "Quiting..."
	close-sqlconnection
	$global:exceldata = $null
	$global:addata = $null
	$global:livedata = $null
	$global:createme =$null
	Exit -1
}

function Main () {
	$Answer = MainPrompt

	switch ($Answer)
	{
		0{process-automated}
		1{process-excel}
		2{process-ad}
		3{process-live}
		4{process-generateemails}
		5{process-quit}
	}
}

#Declairs
[HashTable]$global:exceldata = @{}
[HashTable]$global:addata = @{}
[HashTable]$global:livedata = @{}
[HashTable]$global:createme = @{}
$global:SQLConnection = $null

#Script
Write-Host -foregroundcolor Green -backgroundcolor Black "AlumniMail Import Script" 
while ($true) {
	Main
}
