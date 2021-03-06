﻿# Samples: https://gallery.technet.microsoft.com/scriptcenter/af687d99-5611-4097-97e4-691fda84ad42
$resource = @{
	required = "O paramentro {0} é requerido"
}
function Get-ExcelData {
	# http://blogs.technet.com/b/pstips/archive/2014/06/02/get-excel-data-without-excel.aspx
    [CmdletBinding(DefaultParameterSetName='Worksheet')]
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [String] $Path,

        [Parameter(Position=1, ParameterSetName='Worksheet')]
        [String] $WorksheetName = 'Sheet1',

        [Parameter(Position=1, ParameterSetName='Query')]
        [String] $Query = 'SELECT * FROM [Sheet1$]'
    )

    switch ($pscmdlet.ParameterSetName) {
        'Worksheet' {
            $Query = 'SELECT * FROM [{0}$]' -f $WorksheetName
            break
        }
        'Query' {
            # Make sure the query is in the correct syntax (e.g. 'SELECT * FROM [SheetName$]')
            $Pattern = '.*from\b\s*(?<Table>\w+).*'
            if($Query -match $Pattern) {
                $Query = $Query -replace $Matches.Table, ('[{0}$]' -f $Matches.Table)
            }
        }
    }

    # Create the scriptblock to run in a job
    $JobCode = {
        Param($Path, $Query)

        # Check if the file is XLS or XLSX 
        if ((Get-Item -Path $Path).Extension -eq 'xls') {
            $Provider = 'Microsoft.Jet.OLEDB.4.0'
            $ExtendedProperties = 'Excel 8.0;HDR=YES;IMEX=1'
        } else {
            $Provider = 'Microsoft.ACE.OLEDB.12.0'
            $ExtendedProperties = 'Excel 12.0;HDR=YES'
        }
        
        # Build the connection string and connection object
        $ConnectionString = 'Provider={0};Data Source={1};Extended Properties="{2}"' -f $Provider, $Path, $ExtendedProperties
        $Connection = New-Object System.Data.OleDb.OleDbConnection $ConnectionString

        try {
            # Open the connection to the file, and fill the datatable
            $Connection.Open()
            $Adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $Query, $Connection
            $DataTable = New-Object System.Data.DataTable
            $Adapter.Fill($DataTable) | Out-Null
        }
        catch {
            # something went wrong :-(
            Write-Error $_.Exception.Message
        }
        finally {
            # Close the connection
            if ($Connection.State -eq 'Open') {
                $Connection.Close()
            }
        }

        # Return the results as an array
        return ,$DataTable
    }

    # Run the code in a 32bit job, since the provider is 32bit only
    $job = Start-Job $JobCode -RunAs32 -ArgumentList $Path, $Query
    $job | Wait-Job | Receive-Job
    Remove-Job $job
} #Get-ExcelData

function Get-SharePointListData {
    [CmdletBinding(DefaultParameterSetName='ListName')]
    Param(
        [Parameter(Mandatory=$true, Position=0, HelpMessage="Insira uma URL válida.")]
        [ValidateScript( { [System.Uri]::IsWellFormedUriString($_, [System.UriKind]::Absolute) } ) ]
        [String] $UrlPath,
        
        [Parameter(Mandatory=$true, Position=1, HelpMessage="Informe o Identificador da lista do SharePoint. EX.: 'D6AC1715-8D1D-47BE-94F7-6AE5233B84DD'")]
        [String] $ListID,

        [Parameter(Position=2, ParameterSetName='ListName')]
        [String] $ListName = 'List1',

        [Parameter(Position=2, ParameterSetName='Query')]
        [String] $Query = 'SELECT * FROM [List1]'
    )

    switch ($pscmdlet.ParameterSetName) {
        'ListName' {
            $Query = 'SELECT * FROM [{0}]' -f $ListName
            break
        }
        'Query' {
            # Make sure the query is in the correct syntax (e.g. 'SELECT * FROM [SheetName$]')
            $Pattern = '.*from\b\s*(?<Table>\w+).*'
            if($Query -match $Pattern) {
                $Query = $Query -replace $Matches.Table, ('[{0}]' -f $Matches.Table)
            }
        }
    }

    # Create the scriptblock to run in a job
    $JobCode = {
        Param($UrlPath, $Query, $ListID)

        $ConnectionString = 'Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=1;RetrieveIds=Yes;DATABASE={0};LIST={1};' -f $UrlPath, $ListID
        Using-O ($Connection = Get-ACEConnection $ConnectionString) {		
			try {
				$Adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $Query, $Connection
				$DataTable = New-Object System.Data.DataTable
				$Adapter.Fill($DataTable) | Out-Null
			}
			catch {
				# something went wrong :-(
				Write-Host "Url: $Url"
				Write-Host "Query: $Query"
				Write-Host "ConnectionString: $ConnectionString"
				Write-Error $_.Exception.Message
			}
			if ($Connection.State -eq 'Open') {
				$Connection.Close()
			}
			# Return the results as an array
			return ,$DataTable
		}
    }
	switch ($env:Processor_Architecture) 
    {
		# Run the code in a 32bit job, since the provider is 32bit only
		'x86' { $job = Start-Job $JobCode -InitializationScript $functions -RunAs32 -ArgumentList $UrlPath, $Query, $ListID }
		'AMD64' { $job = Start-Job $JobCode -InitializationScript $functions -ArgumentList $UrlPath, $Query, $ListID }
	}
	$job | Wait-Job | Receive-Job
    Remove-Job $job
} #Get-SharePointListData


$functions = {
	function Using-O {
		# http://weblogs.asp.net/adweigert/powershell-adding-the-using-statement
		[CmdletBinding()]
		param (
			[System.IDisposable] $inputObject = $(throw $resource.required -f "-inputObject"),
			[ScriptBlock] $scriptBlock = $(throw $resource.required -f "-scriptBlock")
		)
    
		Try {
			&$scriptBlock
		} Finally {
			if ($inputObject -ne $null) {
				if ($inputObject.psbase -eq $null) {
					$inputObject.Dispose()
				} else {
					$inputObject.psbase.Dispose()
				}
			}
		}
	} #Using-O

	function Get-ACEConnection {
		param ($ConnectionString)
	
		$conn = New-Object System.Data.OleDb.OleDbConnection $ConnectionString
		try {
			$conn.Open()
			return $conn
		}
		catch {
			installMicrosoftACEOLEDBProvider
		}
		$conn.Open()
		$conn
	} #Get-ACEConnection

	function checkMSAceOledbExist {
		$ie = $null
		try {
			$ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=1;RetrieveIds=Yes;"
			$ie = New-Object System.Data.OleDb.OleDbConnection $ConnectionString
			$ie.Open()
		}
		catch {
			$ie = $null
			Write-Warning $_
		}
		return ($ie -ne $null)
	} #checkMSAceOledbExist

	function installMicrosoftACEOLEDBProvider {
		switch ($env:Processor_Architecture) 
		{ 
			'x86' { $link = 'http://download.microsoft.com/download/2/4/3/24375141-E08D-4803-AB0E-10F2E3A07AAA/AccessDatabaseEngine.exe'; $fileName = "AccessDatabaseEngine.exe"}
			'AMD64' { $link = 'http://download.microsoft.com/download/2/4/3/24375141-E08D-4803-AB0E-10F2E3A07AAA/AccessDatabaseEngine_x64.exe'; $fileName = "AccessDatabaseEngine_x64.exe"}
		}
		$file = "{0}\{1}" -f $env:TEMP, $fileName
		if (-not (Test-Path $file)) {
			$downloader = new-object System.Net.WebClient
			$downloader.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials;
			try {
				Write-Host "INFO: Baixando '$fileName'. Arquiterura $env:Processor_Architecture. Link: $link" -f Cyan
				$downloader.DownloadFile($link, $file)
			} catch {
				Write-Warning "Não foi possível fazer download do componente '$fileName'"
				throw $_.Exception
			}
		}
		Start-Process $file -Wait
	} #installMicrosoftACEOLEDBProvider
}