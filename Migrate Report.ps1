<#




#>

#Parameter
$temp_path_root = "C:\pbimigration\"

# PART 1: Authentication
# ==================================================================
try {
    Get-PowerBIAccessToken 
}
catch {
    Login-PowerBIServiceAccount
}

$tokenhash = Get-PowerBIAccessToken
$token = ($tokenhash.Authorization).Replace("Bearer ", "")


# PART 2: Prompt for user input
# ==================================================================
# Get the list of groups that user is a member of
$all_groups = (Invoke-PowerBIRestMethod -Url 'groups/' -Method Get | ConvertFrom-Json).value 
Clear-Host
Write-Host "Workspaces you have access to:"
""
Write-Host -Object ($all_groups.name -join "`n") -ForegroundColor Yellow 
""


# Ask for the source workspace name
$source_group_ID = ""
while (!$source_group_ID) {
    $source_group_name = Read-Host -Prompt "Enter source workspace"

    if($source_group_name -eq "My Workspace") {
        $source_group_ID = "me"
        break
    }

    Foreach ($group in $all_groups) {
        if ($group.name -eq $source_group_name) {
            if ($group.isReadOnly -eq "True") {
                "Invalid choice: you must have edit access to the group"
                break
            } else {
                $source_group_ID = $group.id
                $source_group_name = $group.name
                break
            }
        }
    }

    if(!$source_group_id) {
        "Please try again, making sure to type the exact name of the group"  
    } 
}

# Ask for source report
$all_reports = Get-PowerBIReport -WorkspaceId $source_group_ID
""
Write-Host "Reports in $source_group_name"
""
Write-Host -Object ($all_reports.name -join "`n") -ForegroundColor Yellow
""

$source_report_ID = ""
while (!$source_report_ID) {
    try {
        $source_report_name = Read-Host -Prompt "Enter source report"
        $source_report_ID = (Get-PowerBIReport -Name $source_report_name -WorkspaceId $source_group_ID).id
    } catch {
        "Did not find a report with that name. Please try again"
        Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
        Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
        continue
    }
}


# Ask for target workspace name
$target_group_ID = "" 
while (!$target_group_id) {
    try {
        $target_group_name = Read-Host -Prompt "Enter target workspace"
        $response = Get-PowerBIWorkspace -Name $target_group_name
        $target_group_id = $response.id
    } catch { 
        "Could not create a group with that name. Please try again and make sure the name is not already taken"
        "More details: "
        Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
        Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
        continue
    }
}

# PART 3: Copying reports and datasets using Export/Import PBIX APIs
# ==================================================================

# Download report from source workspace
"=== Exporting $source_report_name with id: $source_report_id to $temp_path"
$temp_path = $temp_path_root + "$source_report_name.pbix"
$url = "groups/$source_group_ID/reports/$source_report_ID/Export"
try {
    Invoke-PowerBIRestMethod -Url $url -Method Get -OutFile $temp_path
    "Export succeeded!"
    ""
} catch { 
    Write-Host "= This report and dataset cannot be copied, skipping."
    continue
}

# Import report into target workspace
try {
        "=== Importing $source_report_name to target workspace"
        $uri = "https://api.powerbi.com/v1.0/myorg/groups/$target_group_id/imports?datasetDisplayName=$source_report_name.pbix&nameConflict=Abort"

        # Here we switch to HttpClient class to help POST the form data for importing PBIX
        $httpClient = New-Object System.Net.Http.Httpclient $httpClientHandler
        $httpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $token);
        $packageFileStream = New-Object System.IO.FileStream @($temp_path, [System.IO.FileMode]::Open)
        
	    $contentDispositionHeaderValue = New-Object System.Net.Http.Headers.ContentDispositionHeaderValue "form-data"
	    $contentDispositionHeaderValue.Name = "file0"
	    $contentDispositionHeaderValue.FileName = "Financial Performance2.pbix"
 
        $streamContent = New-Object System.Net.Http.StreamContent $packageFileStream
        $streamContent.Headers.ContentDisposition = $contentDispositionHeaderValue
        
        $content = New-Object System.Net.Http.MultipartFormDataContent
        $content.Add($streamContent)

	    $response = $httpClient.PostAsync($uri, $content).Result
 
	    if (!$response.IsSuccessStatusCode) {
		    $responseBody = $response.Content.ReadAsStringAsync().Result
            "= This report cannot be imported to target workspace. Skipping..."
			$errorMessage = "Status code {0}. Reason {1}. Server reported the following message: {2}." -f $response.StatusCode, $response.ReasonPhrase, $responseBody
			throw [System.Net.Http.HttpRequestException] $errorMessage
		} 

        
        # save the import IDs
        $import_job_id = (ConvertFrom-JSON($response.Content.ReadAsStringAsync().Result)).id

        # wait for import to complete
        $upload_in_progress = $true
        while($upload_in_progress) {

            $uri = "https://api.powerbi.com/v1.0/myorg/groups/$target_group_id/imports/$import_job_id"
            $response = Invoke-PowerBIRestMethod -Url $uri -Method Get | ConvertFrom-Json
            
            if ($response.importState -eq "Succeeded") {
                "Publish succeeded!"
                # update the report and dataset mappings
                $report_id = $response.reports[0].id
                $dataset_id = $response.datasets[0].id
                break
            }

            if ($response.importState -ne "Publishing") {
                "Error: publishing failed, skipping this. More details: "
                $response
                break
            }
            
            Write-Host -NoNewLine "."
            Start-Sleep -s 2
        }
            
        
    } catch [Exception] {
        Write-Host $_.Exception
	    Write-Host "== Error: failed to import PBIX"
        Write-Host "= HTTP Status Code:" $_.Exception.Response.StatusCode.value__ 
        Write-Host "= HTTP Status Description:" $_.Exception.Response.StatusDescription
        continue
    }


