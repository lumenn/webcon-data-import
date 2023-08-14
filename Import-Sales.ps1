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
    SELECT DISTINCT
        CONCAT(WFD_ID, '#', WFD_Signature) AS Customer, /*Customer*/
        SalesOrderNumber
    FROM 
        AdventureWorksDW2019.dbo.FactInternetSales JOIN
        BPS_Content.dbo.WFElements ON CustomerKey = WFD_AttText4 /*Customer Key*/ AND WFD_DTYPEID = 1003 /*Customer*/
"@ -ServerInstance "localhost\SQLEXPRESS" -Verbose -TrustServerCertificate

$lumennBuisnessEntityGUID = 'E554D815-F958-463A-B4DD-E2EB29B29FF2'
$saleWorkflowGUID = '5525a5f2-a71b-43eb-914d-1d489ca81da5'
$saleFormGUID = '58a2b9af-d569-4fd6-b8cf-d61528a2085f'
$pathGUID = '9ee1e141-d289-4b75-9785-3b8fc9789a09'

$formFieldGUIDs = @{
    customer = 'dd8a11db-e99b-4b8f-ba5a-d4e772235ca1';
    orderNumber = 'b4a1f00a-be47-45a7-8332-81fbde8b5f71';
    orderedItems = @{
        guid = '0175b3a8-0e1d-4f35-b9ff-57e25d6367bf';
        product = '5402c7e5-338e-4ef0-989e-317a6bf537d6';
        quantity = 'c120eb08-4b50-4692-94bb-37dd7c969e24';
        unitPrice = '480df9a1-4309-42c7-aeb3-a5d8380290e8';
    }
}

$databaseId = 1
$apiVersion = "v5.0"

$i = 1
$errors = [System.Collections.ArrayList]::new()

foreach($row in $results) {
    $saleItems = Invoke-SqlCmd -Query @"
    SELECT
        SalesOrderLineNumber,
        CONCAT(WFD_ID, '#', WFD_Signature) AS Product,
        OrderQuantity AS Quantity,
        UnitPrice AS Price
    FROM
        AdventureWorksDW2019.dbo.FactInternetSales JOIN
        BPS_Content.dbo.WFElements ON ProductKey = WFD_AttText7 /*Product Key*/ AND WFD_DTYPEID = 2004 /*Product form*/
    WHERE
        SalesOrderNumber = '$($row.SalesOrderNumber)' COLLATE DATABASE_DEFAULT
    ORDER BY
        SalesOrderLineNumber
"@ -ServerInstance "localhost\SQLEXPRESS" -Verbose -TrustServerCertificate
    
    $rows = [System.Collections.ArrayList]::new()
    
    foreach($sale in $saleItems) {
        $cells = @(
            @{
                guid = $formFieldGUIDs.orderedItems.product;
                svalue = $sale.Product;
            },
            @{
                guid = $formFieldGUIDs.orderedItems.quantity;
                value = $sale.Quantity;
            },
            @{
                guid = $formFieldGUIDs.orderedItems.unitPrice;
                value = $sale.Price;
            }
        )
        $rows.Add(@{cells = $cells}) > $null
    }

    $listItemsBody = @(
        @{
            guid = $formFieldGUIDs.orderedItems.guid;
            rows = $rows;
        }
    )

    $body = @{
            workflow = @{
                guid = $saleWorkflowGUID
            };
            formType = @{
                guid = $saleFormGUID
            };
            formFields = @(
                @{
                    guid = $formFieldGUIDs.customer;
                    svalue = $row.Customer;
                },
                @{
                    guid = $formFieldGUIDs.orderNumber;
                    value = $row.SalesOrderNumber;
                }
            );
            itemLists = $listItemsBody;
            businessEntity = @{
                guid = $lumennBuisnessEntityGUID;
            }
    }
    $bodyJSON = ConvertTo-Json $body -Depth 10
    try {
        $response = Invoke-RestMethod `
                    -Method Post `
                    -Uri "http://bps.lumenn.local/api/data/$apiVersion/db/$databaseId/elements?path=$pathGUID" `
                    -Body $bodyJSON `
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
    $errors | Export-Csv -Path "$env:USERPROFILE\Downloads\SaleErrors.csv"
}

