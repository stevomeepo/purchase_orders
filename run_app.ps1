# Define an array of directory and script pairs
$scripts = @(
    @{Directory = "C:\Users\allen\Desktop\VS_Code\Inventory_description"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\Marketing_ASIN_Links_Replacement"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\Newproductstatus"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\Purchase_Report"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\Purchase_Orders_to_Print"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\Receiving_Report_V3"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\US_State_Extraction"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\InvoiceChecking_MTLC"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\InvoiceChecking_HJ"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\Enerlites_data_portal"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\Commission_automation"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\EnerlitesPIM\backend"; Script = "app.py"},
    @{Directory = "C:\Users\allen\Desktop\VS_Code\EnerlitesPIM"; Script = "Pricing_Warning.py"}
)

foreach ($item in $scripts) {
    $directory = $item.Directory
    $script = $item.Script

    # Change to the directory
    Set-Location -Path $directory

    # Start the Python script in a separate process (background job)
    Start-Process python -ArgumentList $script -NoNewWindow
}

# Change to the frontend_ui directory
Set-Location -Path "C:\Users\allen\Desktop\VS_Code\EnerlitesPIM\frontend_ui"

# Start the React app using npm
Start-Job -ScriptBlock {
    $npmPath = "C:\Program Files\nodejs\npm.cmd"
    Set-Location -Path "C:\Users\allen\Desktop\VS_Code\EnerlitesPIM\frontend_ui"
    Start-Process $npmPath -ArgumentList "start" -NoNewWindow
}




#  Run C:\Users\allen\Desktop\VS_Code\run_app.ps1