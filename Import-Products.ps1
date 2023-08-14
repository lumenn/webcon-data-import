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
        ProductKey,
        ProductAlternateKey,
        EnglishProductName,
        ISNULL(ListPrice, 0) AS ListPrice,
        CASE 
            WHEN Status IS NULL THEN 0
            WHEN Status = 'Current' THEN 1
        END AS Active
    FROM 
        AdventureWorksDW2019.dbo.DimProduct
"@ -ServerInstance "localhost\SQLEXPRESS" -Verbose -TrustServerCertificate

$lumennBuisnessEntityGUID = 'E554D815-F958-463A-B4DD-E2EB29B29FF2'
$productWorkflowGUID = '2660ca16-457d-432f-8b43-beb282ab999a'
$productFormGUID = '3d9819ff-573a-4d1a-b424-45652e963079'
$pathActiveGUID = 'c6a440c1-51ce-4aa4-a2f3-39cb691f2e88'
$pathBlockedGUID = 'abc2a33f-5bd3-4ba2-85d4-2b9aae166ae2'

$formFieldGUIDs = @{
    name = '7ffc9b32-ad57-4939-af60-d1ab29f6c01c';
    price = '669e369c-1546-4560-9794-44d46b697416';
    erpID = '673a9f06-055f-40b9-b6b6-57b4158db863';
    productKey = '8b54d6c3-a340-4909-b435-f62cf0004eb7';
}

$databaseId = 1
$apiVersion = "v5.0"

$i = 1
$errors = New-Object System.Collections.Generic.List[System.Object]
foreach($row in $results) {
    $requestBody = @{
        workflow = @{
            guid = $productWorkflowGUID;
        }
        formType = @{
            guid = $productFormGUID;
        }
        formFields = @(
            @{
                guid = $formFieldGUIDs.name;
                value = $row.EnglishProductName;
            },
            @{
                guid = $formFieldGUIDs.price;
                value = $row.ListPrice;
            },
            @{
                guid = $formFieldGUIDs.erpID;
                value = $row.ProductAlternateKey;
            },
            @{
                guid = $formFieldGUIDs.productKey;
                value = $row.ProductKey;
            }
        )
        businessEntity = @{
            guid = $lumennBuisnessEntityGUID
        }
    };

    $body = ConvertTo-Json $requestBody -Depth 10

    try {
        $pathGUID = If ($row.Active) {$pathActiveGUID} Else {$pathBlockedGUID}
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


if ($errors.Count -gt 0) {
    $errors | Export-Csv -Path "$env:USERPROFILE\Downloads\ProductErrors.csv"
}

