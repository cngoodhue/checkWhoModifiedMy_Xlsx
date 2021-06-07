Add-Type -AssemblyName System.IO.Compression.FileSystem

# the function below handles the unzipping of whatever file is passed into it

function unzip {
    param(
        [string]$zipFile,
        [string]$outpath
    )

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile, $outpath) *>$null
}

# this function does everything else, basically parsing the inner xml file that resides inside of the excel doc for specific properties

function getProperties {
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter(
            Position=0,
            Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$docPath,                          # parameters
        [Parameter(
            Position=1)]
            [ValidateNotNullOrEmpty()]
            [string]$properties = 'dcterms:modified; cp:lastModifiedBy;' # this is just a string that we are going to parse and use later
    )

    $docPathContent = gci $docPath -Recurse | ? {$_.Extension -eq '.xlsx' -and $_.LastWriteTime -ge '5/1/2021' -and $_.LastWriteTime -le '5/27/2021'} # 'docPath' is the parameter that is passed when the function gets called

    # below is where the magic happens

    foreach ($file in $docPathContent) {
        $properties = $properties.Replace(' ', '') # parsing that string from earlier
        $zipDirName = $env:TEMP + '\' + 'zipTEMP' # creating a temporary directory to store the zip files
	
	# basically if the directory exists, delete it and create a new one
	if (Test-Path $zipDirName) {
        	Remove-Item $zipDirName -Force -Recurse -ErrorAction Ignore | Out-Null
        	New-Item -ItemType Directory -Path $zipDirName | Out-Null
	} else {
		New-Item -ItemType Directory -Path $zipDirName | Out-Null
	}

        unzip $file.FullName $zipDirName # unzip the file and extract it into the directory we just created

        $coreXmlName = $zipDirName + '\docProps\core.xml'
        [xml]$coreXml = Get-Content -Path $coreXmlName

        $reqProperties = $properties.Split(';')
        $docProperties = @{} # instantiate an object to store our stuff in

        foreach ($reqProperty in $reqProperties) {
            $localName = $reqProperty.Split(':')
            $localName = $localName[1]                                                                     # PARSING!!!
            $node = $coreXml.coreProperties.SelectSingleNode("*[local-name(.) = '$localName']")

            $docProperties.Add($reqProperty, $node.innerText)
        }

        Remove-Item $zipDirName -Force -Recurse # remove the temp directory

        Write-Host -ForegroundColor Green "`nFile Name: $($file.Name)`nFile Path: $($file.FullName)`n"             # output
        
        $docProperties 
        Write-Host "`n***************************************************************"
    }
}

$input = Read-Host 'Enter path'

getProperties $input