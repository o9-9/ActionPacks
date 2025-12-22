#Requires -Version 5.0

<#
    .SYNOPSIS
        Sets Automatic Replies for a mailbox
    
    .DESCRIPTION  

    .NOTES
        This PowerShell script was developed and optimized for ScriptRunner. The use of the scripts requires ScriptRunner. 
        The customer or user is authorized to copy the script from the repository and use them in ScriptRunner. 
        The terms of use for ScriptRunner do not apply to this script. In particular, ScriptRunner Software GmbH assumes no liability for the function, 
        the use and the consequences of the use of this freely available script.
        PowerShell is a product of Microsoft Corporation. ScriptRunner is a product of ScriptRunner Software GmbH.
        © ScriptRunner Software GmbH

    .COMPONENT
        Graph target required
        Modules Az.Accounts and Az.Resources are required for Azure data.

    .PARAMETER CsvPath
        [sr-en] Path of the CSV file
        [sr-de] Pfad für die Csv-Datei

    .PARAMETER Delimiter
        [sr-en] Delimiter that separates the property values in the CSV file
        [sr-de] Csv-Trennzeichen

    .PARAMETER FileEncoding
        [sr-en] Type of character encoding that was used in the CSV file
        [sr-de] Encoding der Csv-Datei
#>

param(
    [string]$CsvPath,
    [string]$CsvDelimiter = ';',
    [ValidateSet('Unicode','UTF7','UTF8','ASCII','UTF32','BigEndianUnicode','Default','OEM')]
    [string]$CsvFileEncoding = 'UTF8'
)

[string[]]$includedRoles = @("Application Administrator","Application Developer","Attribute Provisioning Administrator",
        "Authentication Administrator",,"Authentication Extensibility Administrator","B2C IEF Keyset Administrator","Billing Administrator",
        "Cloud Application Administrator","Cloud Device Administrator","Compliance Administrator","Conditional Access Administrator",
        "Directory Writers","Domain Name Administrator","Exchange Administrator" ,"External Identity Provider Administrator",
        "Global Administrator","Global Reader","Helpdesk Administrator","Hybrid Identity Administrator","Intune Administrator",
        "Lifecycle Workflows Administrator","Password Administrator","Privileged Authentication Administrator","Privileged Role Administrator",
        "Security Administrator","Security Operator","Security Reader","SharePoint Administrator","Teams Administrator","User Administrator")

try {
    [string]$resultHtml = @"
    <!DOCTYPE html>
	<html><head>
		<title>Privileged Users Report</title>
		<style>
			body{font-family:'Segoe UI',Arial,sans-serif;margin:20px;background-color:#f5f5f5;color:#333;}
			.header{background-color:#00497d;color:white;padding:20px;border-radius:5px;margin-bottom:20px;}
			.summary{background-color:white;padding:15px;border-radius:5px;margin-bottom:20px;box-shadow:0 2px 4px rgba(0,0,0,0.1);}
			.summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:15px;}
			.summary-item{text-align:center;padding:10px;background-color:#f8f9fa;border-radius:3px;}
			.summary-number{font-size:24px;font-weight:bold;color:#0078d4;}.summary-label{font-size:12px;color:#666;margin-top:5px;}
			.azure-summary{color:#ff6600;}.phishing-resistant{color:#28a745;}`n"
			<!-- Add search bar styles -->
			.search-container{background-color:white;padding:15px;border-radius:5px;margin-bottom:20px;box-shadow:0 2px 4px rgba(0,0,0,0.1);}
			.search-box{width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;font-size:14px;box-sizing:border-box;}
			.search-box:focus{outline:none;border-color:#0078d4;box-shadow:0 0 0 2px rgba(0,120,212,0.1);}
			.search-info{margin-top:10px;font-size:12px;color:#666;}
			table{width:100%;border-collapse:collapse;background-color:white;border-radius:5px;overflow:hidden;box-shadow:0 2px 4px rgba(0,0,0,0.1);font-size:13px;}
			th{background-color:#f8f9fa;padding:10px 8px;text-align:left;font-weight:600;border-bottom:2px solid #dee2e6;cursor:pointer;user-select:none;}
			th:hover{background-color:#e9ecef;}
			td{padding:8px;border-bottom:1px solid #dee2e6;vertical-align:top;}
			tr:hover{background-color:#f8f9fa;}
			tr.hidden{display:none;}
			.status-active{color:#28a745;font-weight:bold;}.status-disabled{color:#dc3545;font-weight:bold;}
			.type-hybrid{color:#fd7e14;font-weight:bold;}.type-cloud{color:#0078d4;font-weight:bold;}
			.mfa-yes{color:#28a745;font-weight:bold;}.mfa-no{color:#dc3545;font-weight:bold;}
			.auth-method-phishing-resistant{color:#28a745;font-weight:bold;}
			.via-group{color:#0078d4;font-weight:bold;}
			.footer{margin-top:20px;text-align:center;color:#666;font-size:12px;}
			.scroll-container{overflow-x:auto;}
			.sortable th::after{content:' ↕';color:#aaa;font-size:10px;}
		</style>
	</head>
    <body>
		<!-- Add header -->
		<div class='header'><h1>Privileged Users Report</h1>
		<p>Generated on: $((Get-Date).ToString('G')) | Signed in as: $($SRXEnv.SRXStartedBy)</p></div>

		<!--  Add summary -->
		<div class='summary'><div class='summary-grid'>
			<div class='summary-item'><div class='summary-number'>%TotalUsers%</div><div class='summary-label'>Total Users</div></div>
			<div class='summary-item'><div class='summary-number'>%ActiveUsers%</div><div class='summary-label'>Active Users</div></div>
			<div class='summary-item'><div class='summary-number'>%HybridUsers%</div><div class='summary-label'>Hybrid Users</div></div>
			<div class='summary-item'><div class='summary-number'>%EntraActiveRolesCount%</div><div class='summary-label'>Entra Active Roles</div></div>
			<div class='summary-item'><div class='summary-number'>%EntraEligibleRolesCount%</div><div class='summary-label'>Entra Eligible Roles</div></div>
		    <div class='summary-item'><div class='summary-number azure-summary'>%AzureActiveRolesCount%</div><div class='summary-label'>Azure Active Roles</div></div>
		    <div class='summary-item'><div class='summary-number azure-summary'>%AzureEligibleRolesCount%</div><div class='summary-label'>Azure Eligible Roles</div></div>
			<div class='summary-item'><div class='summary-number'>%UsersWithoutMFA%</div><div class='summary-label'>Without MFA</div></div>
			<div class='summary-item'><div class='summary-number phishing-resistant'>%UsersWithPhishingResistant%</div><div class='summary-label'>Phishing-Resistant Auth</div></div>
		</div></div>
	
		<!--  Add search bar -->
		<div class='search-container'>
			<input type='text' id='searchBox' class='search-box' placeholder='Search ...'>
			<div class='search-info' id='searchInfo'>Showing all %TotalUsers% users</div>
		</div>

		<!--  Add table -->
		<div class='scroll-container'>
			<table class='sortable' id='dataTable'>
			<thead><tr>`
				<th onclick='sortTable(0)'>UPN</th>
				<th onclick='sortTable(1)'>Entra Active Roles</th>
				<th onclick='sortTable(2)'>Entra Eligible Roles</th>
                <th onclick='sortTable(3)'>Azure Active Roles</th>
				<th onclick='sortTable(4)'>Azure Eligible Roles</th>
				<th onclick='sortTable(5)'>Total</th>
				<th onclick='sortTable(6)'>Status</th>
				<th onclick='sortTable(7)'>Type</th>
				<th onclick='sortTable(8)'>MFA</th>
				<th onclick='sortTable(9)'>Last Interactive Sign In</th>
				<th onclick='sortTable(10)'>Last Non-Interactive Sign In</th>
				<th onclick='sortTable(11)'>Auth Methods</th>
			</tr></thead>	
			<tbody>			
			%USERS_HTML%
			</tbody></table></div>
        <script>
            let currentSortColumn=-1;let sortAscending=true;

            <!-- Sort table function -->
            function sortTable(columnIndex){
                const table=document.getElementById('dataTable');
                const tbody=table.getElementsByTagName('tbody')[0];
                const rows=Array.from(tbody.getElementsByTagName('tr')).filter(row=>!row.classList.contains('hidden'));
                
                if(currentSortColumn===columnIndex){
                    sortAscending=!sortAscending;
                }else{
                    sortAscending=true;
                    currentSortColumn=columnIndex;
                }
                
                rows.sort((a,b)=>{
                    const aText=a.cells[columnIndex].textContent.trim();
                    const bText=b.cells[columnIndex].textContent.trim();
                    const aNum=parseFloat(aText);
                    const bNum=parseFloat(bText);
                    
                    if(!isNaN(aNum)&&!isNaN(bNum)){
                        return sortAscending?aNum-bNum:bNum-aNum;
                    }else{
                        return sortAscending?aText.localeCompare(bText):bText.localeCompare(aText);
                    }
                });
                
                // Re-append all rows (including hidden ones) in sorted order
                const allRows=Array.from(tbody.getElementsByTagName('tr'));
                allRows.forEach(row=>{
                    if(!row.classList.contains('hidden')){
                        tbody.removeChild(row);
                    }
                });
                rows.forEach(row=>tbody.appendChild(row));
            }

            <!-- Search functionality -->
            document.getElementById('searchBox').addEventListener('input', function(e) {
                const searchTerm = e.target.value.toLowerCase();
                const table = document.getElementById('dataTable');
                const tbody = table.getElementsByTagName('tbody')[0];
                const rows = tbody.getElementsByTagName('tr');
                let visibleCount = 0;
                
                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i];
                    const text = row.textContent.toLowerCase();
                    
                    if (searchTerm === '' || text.includes(searchTerm)) {
                        row.classList.remove('hidden');
                        visibleCount++;
                    } else {
                        row.classList.add('hidden');
                    }
                }
                
                // Update search info
                const searchInfo = document.getElementById('searchInfo');
                if (searchTerm === '') {
                    searchInfo.textContent = 'Showing all ' + rows.length + ' users';
                } else {
                    searchInfo.textContent = 'Showing ' + visibleCount + ' of ' + rows.length + ' users';
                }
            });	
        </script>
    </body></html>
"@
#region functions
    function CreateUserHtmlRow{
        <#
            .SYNOPSIS 
                Creates Html row for user
            .PARAMETER UserHtml
                Reference parameter to append user Html
            .PARAMETER UserObject
                User object
        #>

        param(
            [ref]$UserHtml,
            [object]$UserObject

        )

        [string]$usrHtml = @" 
            <tr>
                <td>%UPN%</td>
                <td>%EntraActive%</td>
                <td>%EntraEligible%</td>
                <td>%AzActiveRoles%</td>
                <td>%AzEligibleRoles%</td>
                <td>%TotalRoles%</td>
                <td class='StatusClass'>%AccountStatus%</td>
                <td class='TypeClass'>%UserType</td>
                <td class='MfaClass'>%MFAEnabled%</td>
                <td>%LastInteractiveSignIn%</td>
                <td>%LastNon-InteractiveSignIn%</td>
                <td>%AuthMethods%</td>
            </tr>
"@
        try{
            if($UserObject.'Account Status' -eq 'Active'){
                $usrHtml = $usrHtml.Replace('StatusClass','status-active') 
            } 
            else{ 
                $usrHtml = $usrHtml.Replace('StatusClass','status-disabled') 
            }
            if($UserObject.'User Type' -eq 'Hybrid'){
                $usrHtml = $usrHtml.Replace('TypeClass','type-hybrid') 
            } 
            else{
                $usrHtml = $usrHtml.Replace('TypeClass','type-cloud') 
            }
            if ($UserObject.'MFA Enabled' -eq 'Yes'){
                $usrHtml = $usrHtml.Replace('MfaClass','mfa-yes')
            } 
            else{
                $usrHtml = $usrHtml.Replace('MfaClass','mfa-no')
            }
            $usrHtml = $usrHtml.Replace('%UPN%',$UserObject.'UPN').Replace('%EntraActive%',$UserObject.'Entra Active Roles').Replace('%AuthMethods%',$UserObject.'Auth Methods')
            $usrHtml = $usrHtml.Replace('%EntraEligible%',$UserObject.'Entra Eligible Roles').Replace('%AzActiveRoles%',$UserObject.'Azure Active Roles').Replace('%AzEligibleRoles%',$UserObject.'Azure Eligible Roles')
            $usrHtml = $usrHtml.Replace('%TotalRoles%',$UserObject.'Total Roles').Replace('%AccountStatus%',$UserObject.'Account Status').Replace('%UserType',$UserObject.'User Type')
            $usrHtml = $usrHtml.Replace('%MFAEnabled%',$UserObject.'MFA Enabled').Replace('%LastInteractiveSignIn%',$UserObject.'Last Interactive Sign In').Replace('%LastNon-InteractiveSignIn%',$UserObject.'Last Non-Interactive Sign In')
            $UserHtml.Value += $usrHtml
        }
        catch{
            throw
        }
    }
    function GetAZDatas{
        <#
            .SYNOPSIS                
                Gets the user azure active and eligible roles
            .PARAMETER UserId
                Id of the user
            .PARAMETER ActiveRoles
                Reference parameter for active roles
            .PARAMETER EligibleRoles
                Reference parameter for eligible roles
        #>

        param(
            [string]$UserId,
            [ref]$ActiveRoles,
            [ref]$EligibleRoles
        )

        try {
            $ActiveRoles.Value = @()
            $EligibleRoles.Value = @()
            # Get all subscriptions the user has access to
            $azSubscriptions = Get-AzSubscription -ErrorAction SilentlyContinue
        
            # Get user's group memberships for group-based checks
            $userGroups = @()
            try {
                $groups = Invoke-MgGraphRequest -Uri ("https://graph.microsoft.com/v1.0/users/{0}/memberOf" -f $UserId) -Method GET #harald.pfirmann@1dwps5.onmicrosoft.com
                foreach ($group in $groups.value) {
                    if ($group.'@odata.type' -eq '#microsoft.graph.group') {
                        $userGroups += $group.Id
                    }
                }
            } 
            catch {
                Write-Output "Failed to get user groups: $_"
            }
            foreach ($azSub in $azSubscriptions) {
                try {
                    $null = Set-AzContext -SubscriptionId $azSub.Id -ErrorAction SilentlyContinue
                    
                    # Get active role assignments for this user in the subscription
                    $roleAssignments = Get-AzRoleAssignment -ObjectId $UserId -ErrorAction SilentlyContinue
                    
                    foreach ($rAssignment in $roleAssignments) {
                        # Include privileged roles
                        $isPrivilegedRole = $rAssignment.RoleDefinitionName -eq "Owner" -or 
                                        $rAssignment.RoleDefinitionName -like "*Contributor" -or
                                        $rAssignment.RoleDefinitionName -eq "Reservations Administrator" -or
                                        $rAssignment.RoleDefinitionName -eq "Role Based Access Control Administrator" -or
                                        $rAssignment.RoleDefinitionName -eq "User Access Administrator"
                        
                        if ($isPrivilegedRole) {
                            $scopeInfo = ""
                            
                            if ($rAssignment.Scope -eq "/subscriptions/$($azSub.Id)") {
                                $scopeInfo = "Sub: $($azSub.Name)"
                            } elseif ($rAssignment.Scope -match "/subscriptions/.+/resourceGroups/([^/]+)$") {
                                $rgName = $matches[1]
                                $scopeInfo = "RG: $rgName (Sub: $($azSub.Name))"
                            } else {
                                $resourceName = ($rAssignment.Scope -split "/")[-1]
                                if ($rAssignment.Scope -match "/subscriptions/.+/resourceGroups/([^/]+)/") {
                                    $rgName = $matches[1]
                                    $scopeInfo = "Resource: $resourceName (RG: $rgName, Sub: $($azSub.Name))"
                                } else {
                                    $scopeInfo = "Resource: $resourceName (Sub: $($azSub.Name))"
                                }
                            }
                            
                            $ActiveRoles.Value += "$($rAssignment.RoleDefinitionName) → $scopeInfo"
                        }
                    }
                    
                    # Check for group-based active assignments
                    foreach ($groupId in $userGroups) {
                        $groupAssignments = Get-AzRoleAssignment -ObjectId $groupId -ErrorAction SilentlyContinue
                        
                        if ($groupAssignments) {
                            foreach ($gAssignment in $groupAssignments) {
                                $isPrivilegedRole = $gAssignment.RoleDefinitionName -eq "Owner" -or 
                                                $gAssignment.RoleDefinitionName -like "*Contributor" -or
                                                $gAssignment.RoleDefinitionName -eq "Reservations Administrator" -or
                                                $gAssignment.RoleDefinitionName -eq "Role Based Access Control Administrator" -or
                                                $gAssignment.RoleDefinitionName -eq "User Access Administrator"
                                
                                if ($isPrivilegedRole) {
                                    $scopeInfo = ""
                                    
                                    if ($gAssignment.Scope -eq "/subscriptions/$($azSub.Id)") {
                                        $scopeInfo = "Sub: $($azSub.Name)"
                                    } elseif ($gAssignment.Scope -match "/subscriptions/.+/resourceGroups/([^/]+)$") {
                                        $rgName = $matches[1]
                                        $scopeInfo = "RG: $rgName (Sub: $($azSub.Name))"
                                    } else {
                                        $resourceName = ($gAssignment.Scope -split "/")[-1]
                                        if ($gAssignment.Scope -match "/subscriptions/.+/resourceGroups/([^/]+)/") {
                                            $rgName = $matches[1]
                                            $scopeInfo = "Resource: $resourceName (RG: $rgName, Sub: $($azSub.Name))"
                                        } else {
                                            $scopeInfo = "Resource: $resourceName (Sub: $($azSub.Name))"
                                        }
                                    }
                                    
                                    $ActiveRoles.Value += "$($gAssignment.RoleDefinitionName) → $scopeInfo (via group)"
                                }
                            }
                        }
                    }
                    
                    # Get Azure PIM eligible assignments
                    try {
                        $azContext = Get-AzContext
                        $azureToken = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate(
                            $azContext.Account, 
                            $azContext.Environment, 
                            $azContext.Tenant.Id, 
                            $null, 
                            "Never", 
                            $null, 
                            "https://management.azure.com/"
                        ).AccessToken
                        
                        if ($azureToken) {
                            $headers = @{
                                'Authorization' = "Bearer $($azureToken)"
                                'Content-Type' = 'application/json'
                            }
                            
                            # Direct Azure PIM user assignments 
                            $pimUri = "https://management.azure.com/subscriptions/$($azSub.Id)/providers/Microsoft.Authorization/roleEligibilityScheduleInstances?api-version=2020-10-01&`$filter=principalId eq '$($UserId)'"
                            $pimResponse = Invoke-RestMethod -Uri $pimUri -Headers $headers -Method GET -ErrorAction SilentlyContinue
                            
                            if ($pimResponse.value) {
                                foreach ($eligibleAssignment in $pimResponse.value) {
                                    # Get role definition name
                                    $roleDefId = $eligibleAssignment.properties.roleDefinitionId
                                    $roleDefUri = "https://management.azure.com$($roleDefId)?api-version=2022-04-01"
                                    $roleDefResponse = Invoke-RestMethod -Uri $roleDefUri -Headers $headers -Method GET -ErrorAction SilentlyContinue
                                    
                                    if ($roleDefResponse) {
                                        $roleDefName = $roleDefResponse.properties.roleName
                                        
                                        # Check if it's a privileged role
                                        $isPrivilegedRole = $roleDefName -eq "Owner" -or 
                                                        $roleDefName -like "*Contributor" -or
                                                        $roleDefName -eq "Reservations Administrator" -or
                                                        $roleDefName -eq "Role Based Access Control Administrator" -or
                                                        $roleDefName -eq "User Access Administrator"
                                        
                                        if ($isPrivilegedRole) {
                                            $scope = $eligibleAssignment.properties.scope
                                            $scopeInfo = ""
                                            
                                            if ($scope -eq "/subscriptions/$($azSub.Id)") {
                                                $scopeInfo = "$roleDefName → Subscription ($($azSub.Name))"
                                            } elseif ($scope -match "/subscriptions/.+/resourceGroups/([^/]+)$") {
                                                $rgName = $matches[1]
                                                $scopeInfo = "$roleDefName → Resource Group ($rgName)"
                                            } else {
                                                $resourceName = ($scope -split "/")[-1]
                                                $scopeInfo = "$roleDefName → Resource ($resourceName)"
                                            }
                                            
                                            if ($EligibleRoles.Value -notcontains $scopeInfo) {
                                                $EligibleRoles.Value += $scopeInfo
                                            }
                                        }
                                    }
                                }
                            }
                            
                            # Group-based eligible assignments
                            foreach ($groupId in $userGroups) {
                                $groupPimUri = "https://management.azure.com/subscriptions/$($azSub.Id)/providers/Microsoft.Authorization/roleEligibilityScheduleInstances?api-version=2020-10-01&`$filter=principalId eq '$groupId'"
                                $groupPimResponse = Invoke-RestMethod -Uri $groupPimUri -Headers $headers -Method GET -ErrorAction SilentlyContinue
                                
                                if ($groupPimResponse.value) {
                                    foreach ($eligibleAssignment in $groupPimResponse.value) {
                                        # Get role definition name
                                        $roleDefId = $eligibleAssignment.properties.roleDefinitionId
                                        $roleDefUri = "https://management.azure.com$($roleDefId)?api-version=2022-04-01"
                                        $roleDefResponse = Invoke-RestMethod -Uri $roleDefUri -Headers $headers -Method GET -ErrorAction SilentlyContinue
                                        
                                        if ($roleDefResponse) {
                                            $roleDefName = $roleDefResponse.properties.roleName
                                            
                                            # Check if it's a privileged role
                                            $isPrivilegedRole = $roleDefName -eq "Owner" -or 
                                                            $roleDefName -like "*Contributor" -or
                                                            $roleDefName -eq "Reservations Administrator" -or
                                                            $roleDefName -eq "Role Based Access Control Administrator" -or
                                                            $roleDefName -eq "User Access Administrator"
                                            
                                            if ($isPrivilegedRole) {
                                                $scope = $eligibleAssignment.properties.scope
                                                $scopeInfo = ""
                                                
                                                if ($scope -eq "/subscriptions/$($azSub.Id)") {
                                                    $scopeInfo = "$roleDefName → Subscription ($($azSub.Name)) (via group)"
                                                } elseif ($scope -match "/subscriptions/.+/resourceGroups/([^/]+)$") {
                                                    $rgName = $matches[1]
                                                    $scopeInfo = "$roleDefName → Resource Group ($rgName) (via group)"
                                                } else {
                                                    $resourceName = ($scope -split "/")[-1]
                                                    $scopeInfo = "$roleDefName → Resource ($resourceName) (via group)"
                                                }
                                                
                                                if ($EligibleRoles.Value -notcontains $scopeInfo) {
                                                    $EligibleRoles.Value += $scopeInfo
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Output "Failed to get PIM eligible roles for subscription $($azSub.Name): $_"
                    }                    
                } 
                catch {
                    Write-Output "Failed to process subscription $($azSub.Name): $_"
                    continue
                }
            }
        }
        catch {
            throw
        }
    }
    function GetUserAuthentication {
        <#
            .SYNOPSIS
                Gets the user authentication infos
            .PARAMETER UserId
                Id of the user
            .PARAMETER UserObject
                Reference parameter of the user object
        #>

        param(
            [string]$UserID,
            [ref]$UserObject
        )
        
        try {            
            $UserObject.Value.'MFA Enabled' = 'No'
            $UserObject.Value.'HasPhishingResistant' = $false
        
            # Use beta endpoint
            $uri = "https://graph.microsoft.com/beta/reports/authenticationMethods/userRegistrationDetails('$userId')"
            $authDetails = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        
            if(($authDetails.methodsRegistered -notcontains 'password') -and ($authDetails.methodsRegistered -notcontains 'email')){
                $UserObject.Value.'MFA Enabled' = 'Yes'
            }
            if(($authDetails.methodsRegistered -contains 'windowsHelloForBusiness') -or `
                 ($authDetails.methodsRegistered -contains 'fido2SecurityKey') -or `
                 ($authDetails.methodsRegistered -contains 'passkeyDeviceBound') -or `
                 ($authDetails.methodsRegistered -contains 'passkeyDeviceBoundAuthenticator') -or `
                 ($authDetails.methodsRegistered -contains 'passkeyDeviceBoundWindowsHello')){
                $UserObject.Value.'HasPhishingResistant' = $true
            }
            if(($null -eq $authDetails.methodsRegistered) -or ($authDetails.methodsRegistered.Count -lt 1)){
                $UserObject.Value.'Auth Methods' = 'None'
            }
            else{
                $UserObject.Value.'Auth Methods' = (($authDetails.methodsRegistered | Select-Object -Unique) -join ',')
            }
        } 
        catch {
            $UserObject.Value.'Auth Methods' = 'Unable to check'
            $UserObject.Value.'MFA Enabled' = 'No'
        }
    }
    function GetEligibleRoles {
        <#
            .SYNOPSIS
                Gets the user eligible roles
            .PARAMETER UserId
                Id of the user
            .PARAMETER EligibleRoles
                Reference parameter for eligible roles
        #>

        param(
            [string]$UserId,
            [ref]$EligibleRoles
        )
    
        try{
            $EligibleRoles.Value = @()
        
            # Direct role assignments for the user
            $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules?`$filter=principalId eq '$($UserId)'&`$expand=roleDefinition"
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        
            if ($response.value) {
                foreach ($item in $response.value) {
                    if(($item.status -eq "Provisioned") -and ($null -ne $item.roleDefinition) -and ([System.String]::IsNullOrWhiteSpace($item.roleDefinition.displayName) -eq $false)){
                        $EligibleRoles.Value += $item.roleDefinition.displayName
                    }
                }
            }
            $EligibleRoles.Value = ($EligibleRoles.Value | Sort-Object -Unique)
        }
        catch {
            throw
        }
    }
    function GetUserRoles{
        <#
            .SYNOPSIS
                Gets the user roles
            .PARAMETER UserId
                Id of the user 
            .PARAMETER RelevantRoles
                Roles
            .PARAMETER UserRoles
                Reference parameter for the user roles
        #>

        param(
            [string]$UserId,
            [object[]]$RelevantRoles,
            [ref]$UserRoles
        )

        try {
            [string]$uri = 'https://graph.microsoft.com/v1.0/directoryRoles/{0}/members'
            $UserRoles.Value = @()
            foreach($role in $RelevantRoles){
                $tmpMembers = Invoke-MgGraphRequest -Method GET -Uri ($uri -f $role.ID)
                if($null -ne ($tmpMembers.value | Where-Object {$_.id -eq $UserId})){
                    $UserRoles.Value += $role.DisplayName
                }
            }
            if($UserRoles.Value.Count -lt 1){
                $UserRoles.Value = @('None')
            }
        }
        catch {
            throw
        }
    }
    function GetUserSignIn{
        <#
            .SYNOPSIS
                Gets the user log in infos
            .PARAMETER CloudUser
                User object
            .PARAMETER UserObject
                Reference parameter of user object
        #>
        
        param(
            [object]$CloudUser,
            [ref]$UserObject
        )
        
        try {
            $UserObject.Value.'Last Interactive Sign In' = "Never"
            $UserObject.Value.'Last Non-Interactive Sign In' = "Never"
            if ($null -ne $CloudUser.SignInActivity) {
                if ($CloudUser.SignInActivity.LastSignInDateTime) {
                    try {
                        $UserObject.Value.'Last Interactive Sign In' = ([DateTime]$CloudUser.SignInActivity.LastSignInDateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    } catch {
                        $UserObject.Value.'Last Interactive Sign In' = "Never"
                    }
                }
            if ($CloudUser.SignInActivity.LastNonInteractiveSignInDateTime) {
                try {
                   $UserObject.Value.'Last Non-Interactive Sign In' = ([DateTime]$CloudUser.SignInActivity.LastNonInteractiveSignInDateTime).ToString("yyyy-MM-dd HH:mm:ss")
                } catch {
                   $UserObject.Value.'Last Non-Interactive Sign In' = "Never"
                }
            }
        }
        } 
        catch {
            return "Unknown"
        }
    }
    function GetUserType{
        <#
            .SYNOPSIS
                Checks is the user a hybrid user
            .PARAMETER CloudUser
                User object
            .PARAMETER UserType
                Reference parameter for user type
        #>
        
        param(
            [object]$CloudUser,
            [ref]$UserType
        )
        
        try {
            $UserType.Value.'User Type' = "Cloud"
            # Check if user is synced from on-premises (hybrid)
            if ($CloudUser.OnPremisesSyncEnabled -eq $true) {
                $UserType.Value.'User Type' = "Hybrid"
            } 
            # Check if user has on-premises attributes (additional hybrid check)
            elseif (([string]::IsNullOrEmpty($CloudUser.OnPremisesSecurityIdentifier) -eq $false) -or 
                    ([string]::IsNullOrEmpty($CloudUser.OnPremisesSamAccountName) -eq $false) -or
                    ([string]::IsNullOrEmpty($CloudUser.OnPremisesUserPrincipalName) -eq $false)) {
                $UserType.Value.'User Type' = "Hybrid"
            }
        } 
        catch {
            $UserType.Value.'User Type' = "Unknown"
        }
    }
    function ReadRoles{
        <#
            .SYNOPSIS
                Read directory roles
            .PARAMETER IncRoles
                Included roles
            .PARAMETER Roles
                Reference parameter for selected roles 
        #>

        param(
            [string[]]$IncRoles,
            [ref]$Roles
        )

        try {
            $tmpRoles = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/directoryRoles' -Method GET
            $Roles.Value = ($tmpRoles.value | Where-Object { $_.DisplayName -and $IncRoles.Contains($_.DisplayName) })
        }
        catch {
            throw
        }
    }
#endregion functions
    [System.Collections.ArrayList]$result = New-Object 'System.Collections.ArrayList'
    [string[]]$usrRoles = @()
    [string[]]$eliRoles = @()
    [object[]]$dirRoles = @()
    [object[]]$azActiveRoles = @()
    [object[]]$azEligibleRoles = @()
    [int]$activeUsers = 0
    [int]$hybridUsers = 0
    [int]$entraActiveRolesCount = 0
    [int]$entraEligibleRolesCount = 0
    [int]$azureActiveRolesCount = 0
    [int]$azureEligibleRolesCount = 0
    [int]$usersWithoutMFA = 0
    [int]$usersWithPhishingResistant = 0
    [bool]$readAzureDatas = $true
    # check modules
    if($null -eq (Get-Module -name 'Az.Accounts' -ListAvailable)){
        $readAzureDatas = $false
        Write-Output "Module Az.Accounts not found. Azure data cannot be collected "
    }
    if($null -eq (Get-Module -Name 'Az.Resources' -ListAvailable)){
        $readAzureDatas = $false
        Write-Output "Module Az.Resources not found. Azure data cannot be collected "
    }
    else{
        Import-Module -Name 'Az.Resources' -Force
    }
    # read roles
    ReadRoles -IncRoles $includedRoles -Roles ([ref]$dirRoles)
    $entraUsers = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/users?$select=Userprincipalname,DisplayName,accountenabled,SignInActivity,OnPremisesSyncEnabled,OnPremisesSecurityIdentifier,OnPremisesSamAccountName,OnPremisesUserPrincipalName'
    [string]$usersHtml = ''
    # read users an datas
    foreach($usr in $entraUsers.Value){
        try {            
            [PSCustomObject]$outUsr = [PSCustomObject]@{
                'UPN' = $usr.userPrincipalName
                'Display name' = $usr.displayName
                'Entra Active Roles' = ''
                'Entra Eligible Roles' = ''
                'Azure Active Roles' = 'None'
                'Azure Eligible Roles' = 'None'
                'Total Roles' = 0
                'Account Status' = "Active"
                'User Type' = 'Cloud'
                'MFA Enabled' = 'No'
                'Last Interactive Sign In' = ''
                'Last Non-Interactive Sign In' = ''
                'Auth Methods' = ''
                'HasPhishingResistant' = $false
                'Is Global Admin' = $false
            }
            GetUserType -CloudUser $usr -UserType ([ref]$outUsr)
            GetUserSignIn -CloudUser $usr -UserObject ([ref]$outUsr)
            GetUserAuthentication -UserID $usr.ID -UserObject ([ref]$outUsr)
            GetUserRoles -UserId $usr.ID -RelevantRoles $dirRoles -UserRoles ([ref]$usrRoles)
            GetEligibleRoles -UserId $usr.ID -EligibleRoles ([ref]$eliRoles)            
            $entraEligibleRolesCount += $eliRoles
            $outUsr.'Entra Eligible Roles' = ($eliRoles -join ',<br>')
            $outUsr.'Entra Active Roles' = ($usrRoles -join ',<br>')
            $outUsr.'Total Roles' = $eliRoles.Count
            if(($usrRoles -contains "Global Administrator") -or ($usrRoles -contains "Global Administrator (via group)")){
                $outUsr.'Is Global Admin' = $true 
            }
            # azure roles
            if($readAzureDatas -eq $true){
                GetAZDatas -UserId $usr.ID -ActiveRoles ([ref]$azActiveRoles) -EligibleRoles ([ref]$azEligibleRoles)
            }
            if ($azActiveRoles.Count -gt 0) {
                $sortedRoles = $azActiveRoles | Sort-Object { if ($_ -like "*Owner*") { "0" } else { "1" + $_ } }
                $outUsr.'Azure Active Roles' = $sortedRoles -join ";<br>"
                $azureActiveRolesCount += $sortedRoles.Count
                $outUsr.'Total Roles' += $sortedRoles.Count
            } 
            if ($azEligibleRoles.Count -gt 0) {
                $sortedRoles = $azEligibleRoles | Sort-Object { if ($_ -like "*Owner*") { "0" } else { "1" + $_ } }
                $outUsr.'Azure Eligible Roles' = ($sortedRoles -join ";<br>")
                $azureEligibleRolesCount += $sortedRoles.Count
                $outUsr.'Total Roles' += $sortedRoles.Count
            }
            # user status
            if ($usr.accountEnabled -eq $false){
                $outUsr.'Account Status' = "Disabled" 
            }
            else{
                $activeUsers++
            }
            $null = $result.Add($outUsr)
            # create user html
            CreateUserHtmlRow -UserObject $outUsr -UserHtml ([ref]$usersHtml)
            # statistic
            if($usr.'User Type' -ne 'Cloud'){
                $hybridUsers++
            }
            if($usr.'MFA Enabled' -eq 'No'){
                $usersWithoutMFA++
            }
            if($usr.'HasPhishingResistant' -eq $true){
                $usersWithPhishingResistant++
            }
            if($usrRoles -notcontains 'None'){
                $outUsr.'Total Roles' += $usrRoles.Count
                $entraActiveRolesCount += $usrRoles.Count
            }
        }
        catch {
            Write-Output "ERROR on read user $($usr.userPrincipalName)"
        }
    }
    # create csv file
    if([System.String]::IsNullOrWhiteSpace($CsvPath) -eq $false){        
        $null = $result | Export-Csv -Path ([System.IO.Path]::Combine($CsvPath,"Get-MGUAuditPrivilegedUser_$((Get-Date).ToString('yyyMMdd-hhmmss')).csv" )) -Delimiter $CsvDelimiter -Encoding $CsvFileEncoding -Force -NoTypeInformation
    }
    $resultHtml = $resultHtml.Replace('%UsersWithoutMFA%',$usersWithoutMFA).Replace('%UsersWithPhishingResistant%',$usersWithPhishingResistant)
    $resultHtml = $resultHtml.Replace('%EntraEligibleRolesCount%',$entraEligibleRolesCount).Replace('%EntraActiveRolesCount%',$entraActiveRolesCount)
    $resultHtml = $resultHtml.Replace('%AzureEligibleRolesCount%',$azureEligibleRolesCount).Replace('%AzureActiveRolesCount%',$azureActiveRolesCount)
    $resultHtml = $resultHtml.Replace('%ActiveUsers%',$activeUsers).Replace('%HybridUsers%',$hybridUsers)
    $resultHtml = $resultHtml.Replace('%USERS_HTML%',$usersHtml).Replace('%TotalUsers%',$result.Count)
    if($null -ne $SRXEnv){
        $SRXEnv.ResultMessage = "Report created"
        $SRXEnv.ResultHtml = $resultHtml
    }
    else{
        Write-Output "Report created"
    }
}
catch {
    throw
}