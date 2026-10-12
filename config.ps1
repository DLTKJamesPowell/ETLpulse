$Environments = [ordered]@{
    PRD = @{
        Label      = "Production (WPPPRD02)"
        EtlSchema  = "WPPPRD02_DWH"
        OracleDSN  = "wppprd02_dwh"
    }
    # Add more environments below:
    # DEV = @{
    #     Label      = "Development"
    #     EtlSchema  = "WPPDEV01_DWH"
    #     OracleDSN  = "wppdev01_dwh"
    # }
    # UAT = @{
    #     Label      = "UAT"
    #     EtlSchema  = "WPPUAT01_DWH"
    #     OracleDSN  = "wppuat01_dwh"
    # }
}

$Config = @{
    # Full path to sqlplus.exe if it is not in PATH
    SqlPlusPath    = "sqlplus"

    # HTTP port the dashboard listens on
    # Access at: http://<servername>:8080/
    Port           = 8080

    # How often the browser auto-refreshes (seconds)
    RefreshSeconds = 60
}
