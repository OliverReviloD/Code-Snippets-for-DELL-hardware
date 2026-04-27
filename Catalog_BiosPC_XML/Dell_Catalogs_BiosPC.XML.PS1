 #   file:///C:/Dell/Dell_Catalogs_for_Updates/BiosPc.xml



cls
#region Customizable_Variables
    $Folder_Root               = "C:\Dell"
    $WorkDir                   = "$Folder_Root\Dell_Catalogs_for_Updates"
    $DownloadLink              = "https://dl.dell.com/catalog"
    $DownloadLink              = "https://downloads.dell.com/catalog"
    
    $RefreshCatalogFromCloud   = $false           # remove already downloaded catalog.cab, remove already expanded XML and download again and expand XML again
                                                  # TRUE -> re-use previous downloaded CAB / previously expanded XML
    
    $CatalogName               = "BiosPC.XML"

#endregion Customizable_Variables



Function Main {
    $URL = "$DownloadLink/$CatalogName"

    # "Last-Modified"   on DELL Server
    [DateTime]$Catalog_LastModified = (Invoke-WebRequest -method "Head" $URL  -UseBasicParsing -Verbose -userAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome) | Select Headers -ExpandProperty Headers)["Last-Modified"] 
    Write-Host "$FctName URL-CabFile:   '$URL' - last modified = $Catalog_LastModified "
    $Server_FileDate = $Catalog_LastModified

    # "Content-Length" ( Size )   on DELL Server
    $Catalog_Size = (Invoke-WebRequest -method "Head" $URL  -UseBasicParsing -Verbose -userAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome) | Select Headers -ExpandProperty Headers)["Content-Length"] 
    $Catalog_Size = [Math]::round(    $Catalog_Size / 1kB,2 ) 
    Write-Host "$FctName URL-CabFile:   '$URL' - size = $Catalog_Size"
    $Server_FileSizeKB = $Catalog_Size 

    $LocalFile_FullPath =  "$WorkDir\$CatalogName"
 
    $DL_success = Download-File_From_Cloud -URL_CloudFile $URL -FullPath_LocalFile $LocalFile_FullPath


    [xml]$XMLContent = Get-XMLContent -CatalogXml_FilePath $LocalFile_FullPath
    $Catalog_DateTime = $XMLContent.Manifest.dateTime
    Write-Host "Catalog_DateTime = '$Catalog_DateTime'" # dateTime="2026-04-21T05:40:26-05:00"     version="26.04.21"
    $Content_DateCreated = $Catalog_DateTime


    $AllNodes            = $XMLContent.SelectNodes('/*/*')   | Where { $_.Name -ne 'Prerequisites' -and $_.name -ne 'Signature' } #  all 'SoftwareComponent(s)'
   
    $AllSwComponents = @()
    ForEach ( $Node in $AllNodes )
        {
        $Title       = $Node.Name.Display.'#cdata-section'
        if ( $Title  -ne $Null ) { $Title  = $Title.Trim() }                              # Dell Precision 3550 and Latitude 5410/5510 System BIOS,1.40.1,1.40.1
        Write-Host $Title

        $packageID   = $Node.packageID

        $path        = $Node.path
        $UpdateFile  = $Path.Split("/")[-1]

        $dellVersion = $Node.dellVersion

        #$dateTime    = $Node.dateTime     # can get ignored, DELL internal date, not related to any deployment scenario
        $releaseDate = $Node.releaseDate
        if ( $releaseDate  -ne $Null ) { $releaseDate  = Get-Date $releaseDate -Format "yyyy-MM-ddTHH:mm" }

               
        $ImportantInfo = $Node.ImportantInfo.Url
         
        $Category    = $Node.Category.Display.'#cdata-section'
        if ( $Category  -ne $Null ) { $Category  = $Category.Trim() }
        
        $Criticality = $Node.Criticality.Display.'#cdata-section'
        if ( $Criticality  -ne $Null ) { $Criticality  = $Criticality.Trim() }
        
        $BrandDisplay         = $Node.SupportedSystems.Brand.Display.'#cdata-section'         #  {Latitude, Precision}
        if ( $BrandDisplay  -ne $Null ) { $BrandDisplay  = $BrandDisplay.Trim() }
        
        $BrandModelDisplay    = $Node.SupportedSystems.Brand.Model.Display.'#cdata-section'   
        if ( $BrandModelDisplay  -ne $Null ) { $BrandModelDisplay  = $BrandModelDisplay.Trim() }
        

        $BrandModelSystemID   = $Node.SupportedSystems.Brand.Model.SystemID
        
        # only deugging
        $CountString = "$($BrandDisplay.Count) -- $($BrandModelDisplay.Count) -- $($BrandModelSystemID.Count )"
        if ( $Title -eq "Dell Precision 3550 and Latitude 5410/5510 System BIOS,1.40.1,1.40.1" ) { $CountString }


        if ( $BrandDisplay -is [array] )
            {
            for($j = 0; $j -lt $BrandDisplay.Count; $j++)
                {
               # Write-Host " $($BrandDisplay.Count) $Title  "
                
                if ($packageID -eq "g8t8n88888")
                    {
                    Write-Host  $BrandDisplay[$j]    -ForegroundColor Cyan
                    Write-Host  $BrandModelDisplay   -ForegroundColor Cyan
                    }

                if ( $BrandModelDisplay.Count -gt 1) 
                    { 
                    
                    for($i = 0; $i -lt $BrandModelDisplay.Count; $i++)
                        {

                        if ($packageID -eq "g8t8n99999999") {      Write-Host ("{0} {1} {2}" -f  $BrandDisplay, $BrandModelDisplay[$i] , $BrandModelSystemID[$i]  ) }
                        if ( $BrandModelDisplay[$i] -match  $BrandDisplay[$j] )
                            {
                            $AllSwComponents += New-Object -TypeName PSObject -Property @{ Category=$Category; Criticality=$Criticality; BrandDisplay=$BrandDisplay[$j]; BrandModelDisplay=$BrandModelDisplay[$i] ; BrandModelSystemID=$BrandModelSystemID[$i]; packageID=$packageID;  Title=$Title ;  UpdateFile=$UpdateFile; dellVersion=$dellVersion; releaseDate=$releaseDate  ;ImportantInfo=$ImportantInfo } 
                            }
                        }
                    }
                else
                    {
                     $AllSwComponents += New-Object -TypeName PSObject -Property @{ Category=$Category; Criticality=$Criticality; BrandDisplay=$BrandDisplay[$j]; BrandModelDisplay=$BrandModelDisplay; BrandModelSystemID=$BrandModelSystemID; packageID=$packageID;  Title=$Title ;  UpdateFile=$UpdateFile; dellVersion=$dellVersion; releaseDate=$releaseDate ;ImportantInfo=$ImportantInfo } 
                    }
                }
            }
        else
            {
            if ( $BrandModelDisplay.Count -gt 1 ) 
                    { 
                    for($i = 0; $i -lt $BrandModelDisplay.Count; $i++)
                        {
                         #     Write-Host ("{0} {1} {2}" -f  $BrandDisplay, $BrandModelDisplay[$i] , $BrandModelSystemID[$i]  )
                        $AllSwComponents += New-Object -TypeName PSObject -Property @{ Category=$Category; Criticality=$Criticality; BrandDisplay=$BrandDisplay; BrandModelDisplay=$BrandModelDisplay[$i] ; BrandModelSystemID=$BrandModelSystemID[$i]; packageID=$packageID;  Title=$Title ;  UpdateFile=$UpdateFile; dellVersion=$dellVersion; releaseDate=$releaseDate ;ImportantInfo=$ImportantInfo } 
                        }
                    }
                else
                    {
                     $AllSwComponents += New-Object -TypeName PSObject -Property @{ Category=$Category; Criticality=$Criticality; BrandDisplay=$BrandDisplay; BrandModelDisplay=$BrandModelDisplay; BrandModelSystemID=$BrandModelSystemID; packageID=$packageID;  Title=$Title ;  UpdateFile=$UpdateFile ; dellVersion=$dellVersion; releaseDate=$releaseDate;ImportantInfo=$ImportantInfo } 
                    }
            }
  
      
        }
        $AllSwComponents | sort BrandDisplay,BrandModelDisplay,BrandModelSystemID | select Category,Criticality,packageID,BrandDisplay,BrandModelDisplay ,BrandModelSystemID ,Title,UpdateFile,dellVersion,releaseDate,ImportantInfo | ft | out-String | Write-Host   # | Select-Object -Last 5

        Write-Host "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww"
        $GridViewTitle = "Dell catalog 'BiosPC.XML' only SINGLE-LINE select is possible - please close with 'OKAY' ( bottom right ) "
        Write-Host "W  - starting GridView / GUI  - $GridViewTitle"
        Write-Host "W------------------------------------------------------------------------------w"
        $Selected = $AllSwComponents  | sort BrandDisplay,BrandModelDisplay,BrandModelSystemID  | select Category,Criticality,packageID,BrandDisplay,BrandModelDisplay ,BrandModelSystemID ,Title,UpdateFile,dellVersion,releaseDate,ImportantInfo | Out-GridView -Title $GridViewTitle  -OutputMode Single
        Write-Host "W------------------------------------------------------------------------------w"
        Write-Host "W  - result of GridView / GUI                                                  w"
         if ( $Selected ) 
            {
            $Selected | out-string | Write-Host 
            }
         else
            {
            return 1 
            }
 }     

Function Get-XMLContent 
    {
    Param ( [String]$CatalogXml_FilePath)
    
    $FctName = "Get-XMLContent() -"
    # ##########################
    If( Test-Path $CatalogXml_FilePath) {  
        $FileSize = [Math]::round(    (get-Item -path $CatalogXml_FilePath).Length / 1MB,2 ) 
        Write-Host "$FctName Import XML file ('$CatalogXml_FilePath', $FileSize MB) and create PS1 XML Object"
        [xml]$XML = Get-Content $CatalogXml_FilePath -Verbose
        }
    else                                           
        {
        Write-Warning "$FctName Import XML file ('$CatalogXml_FilePath') is missing" 
        [xml]$XML = $NULL
        }           
    
    return $XML
    }


Function Download-File_From_Cloud 
    {
    Param ( [String]$URL_CloudFile,[String]$FullPath_LocalFile)
   $FctName = "Download-File_From_Cloud() -"
   if (Test-Path $FullPath_LocalFile) {  Remove-Item   -Path $FullPath_LocalFile  -Force -Verbose}
   
    Write-Host "$FctName Downloading file from URL"
    Write-Host "$FctName     from cloud '$URL_CloudFile'"
    Write-Host "$FctName     to local   '$FullPath_LocalFile'"
    Invoke-WebRequest -Uri $URL_CloudFile -OutFile $FullPath_LocalFile -UseBasicParsing -userAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome) -Verbose
    [int32]$n=1
    While(!(Test-Path $n) -and $n -lt '3'){
      #  Invoke-WebRequest -Uri $URL_CloudFile -OutFile $FullPath_LocalFile -UseBasicParsing -Verbose
        Invoke-WebRequest -Uri $URL_CloudFile -OutFile $FullPath_LocalFile -UseBasicParsing -userAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome)
        <#
            Invoke-WebRequest -Uri $UriLink -OutFile $TargetFilePath -UseBasicParsing -Proxy $ProxyServer
            
                if download started without  '-userAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome)'  
                -> DELL server will refuse access / download
                -> System.Net.WebException: The remote server returned an error: (403) Forbidden.

                You don't have permission to access "https://dl.dell.com/FOLDER12702584M/3/Dell-ControlVault3-Driver-and-Firmware_G7K77_WIN64_5.15.10.14_A31_01.EXE" on this server.

            Solution:
            ----------- 
            Invoke-WebRequest -Uri $UriLink -OutFile $TargetFilePath -UseBasicParsing -userAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome)
        #>
        $n++
        }
    if (Test-Path $FullPath_LocalFile)   
        {
        Write-Host "$FctName File is downloaded to '$FullPath_LocalFile'"
        return $True
        }
    return $False
    }

 return Main
