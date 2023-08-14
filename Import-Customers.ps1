Import-Module sqlserver

$clientId = "a8494215-ace2-4c77-b84c-f13cbe1b6c2b"
$clientSecret = "ifUCBb/qU8v88m00sQtQ9HRCW43xPuaQ1/ivozFmCt0="
$scopes = "App.Elements.ReadWrite.All"

$authorization = Invoke-RestMethod `
    -Method Post `
    -Uri "http://bps.lumenn.local/api/oauth2/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body "grant_type=client_credentials&client_id=$clientId&client_secret=$clientSecret&scope=$scopes"

$token = $authorization.access_token

$results = Invoke-SqlCmd -Query @"
    SELECT
        CustomerKey,
        FirstName,
        MiddleName,
        LastName,
        BirthDate,
        CASE
        WHEN Gender = 'M' THEN 'Male'
        WHEN Gender = 'F' THEN 'Female'
        END AS Gender
    FROM 
        AdventureWorksDW2019.dbo.DimCustomer
"@ -ServerInstance "localhost\SQLEXPRESS" -Verbose -TrustServerCertificate

$lumennBuisnessEntityGUID = 'E554D815-F958-463A-B4DD-E2EB29B29FF2'
$customerWorkflowGUID = '8a5d448e-f8cd-45ff-a7fd-b3138390d32b'
$customerFormGUID = '3017e844-6313-4ce2-ace0-ae38d913b77b'
$pathGUID = 'a7465986-c850-4222-ae87-d07bb356004c'

$formFieldGUIDs = @{
    firstName = '34a16cc2-eb87-4d86-be05-62f9262cb79e';
    middleName = '1d6dd53d-89d8-4bc3-8f74-fef4f0ad3cd1';
    lastName = '7bc907b6-75ce-4f54-9453-e1ff526505b5';
    birthDate = '70673133-bfe7-452a-981f-4b6d0b2e16db';
    gender = 'b5e53681-e96d-4596-a2c8-260546882ffe';
    customerKey = 'd539357f-6ae7-4d0c-a05a-ba5f3080a650';
}

$databaseId = 1
$apiVersion = "v5.0"

$i = 1
$errors = New-Object System.Collections.Generic.List[System.Object]

foreach($row in $results) {
    $requestBody = @{
            workflow = @{
                guid = "$customerWorkflowGUID"
            }
            formType = @{
                guid = "$customerFormGUID"
            }
            formFields = @(
                @{
                    guid = $formFieldGUIDs.firstName;
                    value = $row.FirstName;
                },
                @{
                    guid = $formFieldGUIDs.middleName;
                    value = $row.MiddleName;
                },
                @{
                    guid = $formFieldGUIDs.lastName;
                    value = $row.LastName;
                },
                @{
                    guid = $formFieldGUIDs.birthDate;
                    value = Get-Date -Date $row.BirthDate -Format "o";
                },
                @{
                    guid = $formFieldGUIDs.gender;
                    value = $row.Gender;
                },
                @{
                    guid = $formFieldGUIDs.customerKey;
                    value = $row.CustomerKey;
                }
            )
            businessEntity = @{
                guid = $lumennBuisnessEntityGUID
            }
    }

    $body = ConvertTo-Json $requestBody -Depth 10

    try {
        $response = Invoke-RestMethod `
                    -Method Post `
                    -Uri "http://bps.lumenn.local/api/data/$apiVersion/db/$databaseId/elements?path=$pathGUID" `
                    -Body $body `
                    -ContentType "application/json" `
                    -Headers @{Authorization = "Bearer $token"}
    }
    catch {
        $errors.Add($row)
    }

    Write-Progress -Activity "Import in progress" -Status "$i out of $($results.Length)"
    $i++;
}

$errors | Export-Csv -Path "$env:USERPROFILE\Downloads\CustomerErrors.csv"
