<#
        Dell Client Commercial PCs Update Catalog Cadences   
            https://downloads.dell.com
            https://dl.dell.com

        https://www.dell.com/support/kbdoc/en-us/000453493/dell-client-commercial-pcs-update-catalog-cadences
        Name	            File Name	                                       Cadence
        ################################################################################################
        Software Gallery	DellAppCatalogPC.cab		                       Monthly Once(Last working day)

        SDP Catalog	        DellSDPCatalogPC.cab	   	                       Monthly twice 2nd and 4th Tuesday

        PDK	                CatalogPC.cab	                                   Quarterly once

        Index	            CatalogIndexPC.cab	                               Every Tuesday and Friday

        Platform	        File Name according to platform System identifier  Every Tuesday and Friday
                            The exact folder is changing when there is an 
                            update in the catalog and is per platform.	
        ================================================================================================
                            
#>


<#
    
        Result from 2026-04-23

        Product                          FileName                        Cadence                           FileSize (KB) Server_FileDate     Content_DateCreated       CatalogAge_Days
        -------                          --------                        -------                           ------------- ---------------     -------------------       ---------------
        DCC                              DellCommandConfigureCatalog.cab unknown                                     674 10.04.2025 07:00:46                            version 2.0.70
        DCU                              CatalogIndexPC.cab              Every Tuesday and Friday                    413 21.04.2026 12:27:39 2026-04-21T03:20:21-05:00               2
        DCU (N-1)                        DTUECatalogIndexPC.cab          Monthly ?                                   410 17.04.2026 16:16:44 2026-04-16T03:30:55-05:00               7
        Enterprise-OSD                   DriverPackCatalog.CAB           unknown                                     319 14.04.2026 13:18:27 2026-04-14T10:56:02                     9
        PDK                              CatalogPC.cab                   Quarterly once                             2564 18.03.2026 07:42:43 2026-03-17T07:42:29-05:00              37
        SCCM-Dell-Business-Client-Update DellSDPCatalogPC.cab            Monthly twice 2nd and 4th Tuesday          9775 10.04.2026 14:14:47 2026-04-09T05:58:50                    14
        Software Gallery                 DellAppCatalogPC.cab            Monthly Once (Last working day)              54 30.03.2026 09:26:29 2026-03-13T03:58:50                    41
        unknown                          CatalogPCIT.cab                 unknown                                     127 12.10.2021 15:24:55 2021-08-06T18:58:14+05:30            1721
   
    
#>




Function Main {

#region Customizable_Variables
    $Folder_Root               = "C:\Dell"
    $WorkDir                   = "$Folder_Root\Dell_Catalogs_for_Updates"
    $DownloadLink              = "https://downloads.dell.com/catalog"
    $DownloadLink              = "https://dl.dell.com/catalog"
    $RefreshCatalogFromCloud   = $True



    $DellCatalogs = @()
    # $DellCatalogs += New-Object -TypeName PSObject -Property @{ Product="Bios";                              FileName="BiosPc.xml";                            Cadence="unknown"                  ;CatalogAge_Days=$Null; Content_DateCreated=$Null; Server_FileDate=$Null; 'FileSize (KB)'=$Null } 
    $DellCatalogs += New-Object -TypeName PSObject -Property @{ Product="DCU";                               FileName="CatalogIndexPC.cab";                    Cadence="Every Tuesday and Friday" ;CatalogAge_Days=$Null; Content_DateCreated=$Null; Server_FileDate=$Null; 'FileSize (KB)'=$Null } 
    $DellCatalogs += New-Object -TypeName PSObject -Property @{ Product="DCU (N-1)";                         FileName="DTUECatalogIndexPC.cab";                Cadence="Monthly ?"                ;CatalogAge_Days=$Null; Content_DateCreated=$Null; Server_FileDate=$Null; 'FileSize (KB)'=$Null } 
    $DellCatalogs += New-Object -TypeName PSObject -Property @{ Product="PDK";                               FileName="CatalogPC.cab";                         Cadence="Quarterly once"           ;CatalogAge_Days=$Null; Content_DateCreated=$Null; Server_FileDate=$Null; 'FileSize (KB)'=$Null } 
    # $DellCatalogs += New-Object -TypeName PSObject -Property @{ Product="unknown";                           FileName="CatalogPCIT.cab";                       Cadence="unknown"                  ;CatalogAge_Days=$Null; Content_DateCreated=$Null; Server_FileDate=$Null; 'FileSize (KB)'=$Null } 
    $DellCatalogs += New-Object -TypeName PSObject -Property @{ Product="Software Gallery";                  FileName="DellAppCatalogPC.cab";                  Cadence="Monthly Once (Last working day)"  ;CatalogAge_Days=$Null; Content_DateCreated=$Null; Server_FileDate=$Null; 'FileSize (KB)'=$Null } 
    $DellCatalogs += New-Object -TypeName PSObject -Property @{ Product="DCC";                               FileName="DellCommandConfigureCatalog.cab";       Cadence="unknown"                  ;CatalogAge_Days=$Null; Content_DateCreated=$Null; Server_FileDate=$Null; 'FileSize (KB)'=$Null } 
    $DellCatalogs += New-Object -TypeName PSObject -Property @{ Product="SCCM-Dell-Business-Client-Update";  FileName="DellSDPCatalogPC.cab";                  Cadence="Monthly twice 2nd and 4th Tuesday" ;CatalogAge_Days=$Null; Content_DateCreated=$Null; Server_FileDate=$Null; 'FileSize (KB)'=$Null } 
    $DellCatalogs += New-Object -TypeName PSObject -Property @{ Product="Enterprise-OSD";                    FileName="DriverPackCatalog.CAB";                 Cadence="unknown"                  ;CatalogAge_Days=$Null; Content_DateCreated=$Null; Server_FileDate=$Null; 'FileSize (KB)'=$Null } 

#endregion Customizable_Variables


    cls
    ForEach ( $Catalog in $DellCatalogs )
        {
        $URL = "$DownloadLink/$($Catalog.FileName)"

        # "Last-Modified"   on DELL Server
        [DateTime]$Catalog_LastModified = (Invoke-WebRequest -method "Head" $URL  -UseBasicParsing -Verbose -userAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome) | Select Headers -ExpandProperty Headers)["Last-Modified"] 
        Write-Host "$FctName URL-CabFile:   '$URL' - last modified = $Catalog_LastModified "
        $Catalog.Server_FileDate = $Catalog_LastModified

        # "Content-Length" ( Size )   on DELL Server
        $Catalog_Size = (Invoke-WebRequest -method "Head" $URL  -UseBasicParsing -Verbose -userAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome) | Select Headers -ExpandProperty Headers)["Content-Length"] 
        $Catalog_Size = [Math]::round(    $Catalog_Size / 1kB,2 ) 
        Write-Host "$FctName URL-CabFile:   '$URL' - size = $Catalog_Size"
        $Catalog.'FileSize (KB)' = $Catalog_Size 
       
        if ( $Catalog.FileName -match ".cab" -and $Catalog.FileName -ne "DellCommandConfigureCatalog.cab")
            {
            $Local_CatalogXMLFileName  = $($Catalog.FileName).Replace(".CAB",".XML").Replace(".cab",".XML")
            $DL_success_expand = Download-Catalog_and_Expand -TargetDir $WorkDir                     `
                        -Force_NewDownload $RefreshCatalogFromCloud             `
                        -URL_CabFile       $URL                                 `
                        -Local_CabFileName $($Catalog.FileName)                 `
                        -FileNameToExpand  $Local_CatalogXMLFileName  
        
            if ( $Catalog.FileName -eq "CatalogPC.cab" -or $Catalog.FileName -eq "CatalogPCIT.cab"  )
                {
                [xml]$XMLContent = Get-XMLContent -CatalogXml_FilePath "$WorkDir\$Local_CatalogXMLFileName" 
                $Catalog_DateTime = $XMLContent.Manifest.DateTime
                Write-Host "Catalog_DateTime = '$Catalog_DateTime'" 
                $Catalog.Content_DateCreated = $Catalog_DateTime
                }
                   
            if ( $Catalog.FileName -eq "DriverPackCatalog.cab"  )
                {
                [xml]$XMLContent = Get-XMLContent -CatalogXml_FilePath "$WorkDir\$Local_CatalogXMLFileName" 
                $Catalog_DateTime = $XMLContent.DriverPackManifest.DateTime
                Write-Host "Catalog_DateTime = '$Catalog_DateTime'" 
                $Catalog.Content_DateCreated = $Catalog_DateTime
                }
        
            if ( $Catalog.FileName -eq "CatalogIndexPC.cab"  -or  $Catalog.FileName -eq "DTUECatalogIndexPC.cab")
                {
                [xml]$XMLContent = Get-XMLContent -CatalogXml_FilePath "$WorkDir\$Local_CatalogXMLFileName" 
                $Catalog_creationDateTime = $XMLContent.ManifestIndex.creationDateTime
                Write-Host "Catalog_creationDateTime = '$Catalog_creationDateTime'" 
                $Catalog.Content_DateCreated = $Catalog_creationDateTime
                }
         
             if ( $Catalog.FileName -eq "DellAppCatalogPC.cab"  -or $Catalog.FileName -eq "DellSDPCatalogPC.cab"   )
                {
                Remove-Variable XMLContent -ErrorAction SilentlyContinue
                [xml]$XMLContent = Get-XMLContent -CatalogXml_FilePath "$WorkDir\$Local_CatalogXMLFileName" 
                Write-Host "- enumerate all software package, search for most recent 'CreationDate'"
                # SystemsManagementCatalog.SoftwareDistributionPackage.Properties - PublicationState="Published" 
                # SystemsManagementCatalog.SoftwareDistributionPackage.Properties - CreationDate="2025-12-11T23:52:28" 
                # $XMLContent.DocumentElement.FirstChild.Properties.CreationDate    # of the first SW package
                # $RootNodes           = $XMLContent.SelectNodes('/*')      #  Namspaces and 'SoftwareDistributionPackage'
                  $AllNodes            = $XMLContent.SelectNodes('/*/*')    #  all 'SoftwareDistributionPackage(s)'
                # $AllNodes_Properties = $XMLContent.SelectNodes('/*/*/*')  #  all 'SoftwareDistributionPackage(s)' - a lot of info - unstructured content
                $AllSwp = @()
                ForEach ( $Node in $AllNodes )
                    {
                    $SwPkgDate  = $Node.Properties.CreationDate
                    $SwPkgTitle = $Node.LocalizedProperties.Title
                    $SwPkgState = $Node.Properties.PublicationState
                    $AllSwp += New-Object -TypeName PSObject -Property @{ SwPkgDate=$SwPkgDate;  SwPkgTitle=$SwPkgTitle ;  SwPkgState=$SwPkgState  } 
                    }
                $AllSwp | sort SwPkgDate | Select-Object -Last 5 | ft | out-String | Write-Host
                Write-Host " .... only last 5 updates get displayed ....." -ForegroundColor Yellow 
                $Catalog_DateTime =  ($AllSwp | sort SwPkgDate -Descending | select SwPkgDate -Unique -First 1).SwPkgDate
                Write-Host "- most recent 'CreationDate' = '$Catalog_DateTime'"
                $Catalog.Content_DateCreated = $Catalog_DateTime
                }
            }
        else
            {
            $LocalFile_FullPath = "$WorkDir\$($Catalog.FileName)"   
            $DL_success = Download-File_From_Cloud -URL_CloudFile $URL -FullPath_LocalFile $LocalFile_FullPath

            if ( $Catalog.FileName -eq "DellCommandConfigureCatalog.cab")
                { 
                Write-Host "expanding 'Version.txt' vom CAB file" 
                $expand_success = Expand-File_From_CAB -FilePath_CAB "$Workdir\$($Catalog.FileName)" -FolderToExpandTo $Workdir -FileNameToExpand "Version.txt"
                If ( Test-Path  "$Workdir\Version.txt" )
                    {
                    $TxtContent = Get-Content "$Workdir\Version.txt"
                    Write-Host "'Version.txt' Content = '$TxtContent'"   #   e.g.  2.0.70
                    $CatalogAge_Days = "version $TxtContent" 
                    Remove-Item -path "$Workdir\Version.txt"
                    }
                else
                    {
                    $CatalogAge_Days = "version unknown" 
                    }
                }

            if ( $Catalog.FileName -eq "BiosPc.xml" )
                {
                [xml]$XMLContent = Get-XMLContent -CatalogXml_FilePath $LocalFile_FullPath
                $Catalog_DateTime = $XMLContent.Manifest.dateTime
                Write-Host "Catalog_DateTime = '$Catalog_DateTime'" # dateTime="2026-04-21T05:40:26-05:00"     version="26.04.21"
                $Catalog.Content_DateCreated = $Catalog_DateTime
                }
            }
        
        # calculate catalog age   - when last time upodated 
        $Today = Get-Date
        $CatalogAge_Days = (New-TimeSpan -Start $Catalog.Content_DateCreated -End $Today -ErrorAction SilentlyContinue).days
        if (   $CatalogAge_Days -ne $Null )  {    
             Write-Host "CatalogAge_Days = '$CatalogAge_Days' days"  
             $Catalog.CatalogAge_Days=$CatalogAge_Days 
             }  
        else 
            {  
            $Catalog.CatalogAge_Days= "version $TxtContent" 
            }
    
        Remove-Variable XMLContent      -ErrorAction SilentlyContinue
        Remove-Variable CatalogAge_Days -ErrorAction SilentlyContinue
        }

    $DellCatalogs | sort Product | select Product,FileName,Cadence,'FileSize (KB)',Server_FileDate,Content_DateCreated, CatalogAge_Days | ft

    }


Function Get-XMLContent 
    {
    Param ( [String]$CatalogXml_FilePath)
    
    $FctName = "Get-XMLContent() -"
    # ##########################
    $FileSize = [Math]::round(    (get-Item -path $CatalogXml_FilePath).Length / 1MB,2 ) 
    Write-Host "$FctName Import XML file ('$CatalogXml_FilePath', $FileSize MB) and create PS1 XML Object"
    If( Test-Path $CatalogXml_FilePath)   {[xml]$XML = Get-Content $CatalogXml_FilePath -Verbose}
    else                                           {[xml]$XML = $NULL}           
    return $XML
    }

Function Download-Catalog_and_Expand
    {
    Param ( [String]$URL_CabFile,[String]$Local_CabFileName, 
            [String]$TargetDir, [String]$FileNameToExpand,
            [bool]$Force_NewDownload)

    $FctName = "Download-Catalog_and_Expand() -"
    Write-Host "$FctName URL_CabFile       = '$URL_CabFile'"
    Write-Host "$FctName Local_CabFileName = '$Local_CabFileName'"
    Write-Host "$FctName TargetDir         = '$TargetDir'"
    Write-Host "$FctName FileNameToExpand  = '$FileNameToExpand'"
    Write-Host "$FctName Force_NewDownload = '$Force_NewDownload'"


   # [DateTime]$URL_CabFile_LastModified = (Invoke-WebRequest -method "Head" $URL_CabFile  -UseBasicParsing -Verbose | Select Headers -ExpandProperty Headers)["Last-Modified"]
    [DateTime]$URL_CabFile_LastModified = (Invoke-WebRequest -method "Head" $URL_CabFile  -UseBasicParsing -Verbose -userAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome) | Select Headers -ExpandProperty Headers)["Last-Modified"] 
    Write-Host "$FctName URL-CabFile:   '$URL_CabFile' - last modified = $URL_CabFile_LastModified "

    If( -not (Test-Path $TargetDir) )                    { $NewFolder = New-Item      -Path $TargetDir -ItemType Directory}

    $Local_CabFile_FullPath      = "$TargetDir\$Local_CabFileName"
    $Local_ExpandedFile_FullPath = "$TargetDir\$FileNameToExpand"

    if (Test-Path $Local_CabFile_FullPath) {  
        Write-Host "$FctName Local-CabFile: '$Local_CabFile_FullPath'   exists"
        [DateTime]$Local_CabFile_LastWriteTime = (get-Item -Path $Local_CabFile_FullPath).LastWriteTime
        Write-Host "$FctName Local-CabFile: '$Local_CabFile_FullPath' - Last WriteTime = $Local_CabFile_LastWriteTime "

        $DelayDays = (New-TimeSpan -Start $URL_CabFile_LastModified -End $Local_CabFile_LastWriteTime).days
        Write-Host "$FctName -- difference between local file date and cloud file date is '$DelayDays' days"

        if ($DelayDays -gt 14 ) { 
            Write-Host "$FctName -- a download of the most recent Catalog will get started"
            $Force_NewDownload = $True 
            }
        }
    else
        {
        Write-Host "$FctName Local-CabFile: '$Local_CabFile_FullPath'   does not exist - download will get forced" -ForegroundColor Yellow
        $Force_NewDownload = $True 
        }

    # --------------------------------------
    If ( $Force_NewDownload -ne $False )
        {
        if (Test-Path $Local_CabFile_FullPath)         {  
            Write-Host "$FctName Local-CabFile:  '$Local_CabFile_FullPath'  will get deleted"
            Remove-Item -Path $Local_CabFile_FullPath       -Force -Verbose
            }
        if (Test-Path $Local_ExpandedFile_FullPath)    {  
            Write-Host "$FctName Local-CabFile:  '$Local_ExpandedFile_FullPath'  will get deleted"
            Remove-Item -Path $Local_ExpandedFile_FullPath  -Force -Verbose
            }
        }
   
    
    # --------------------------------------
    $DL_success = $True
    if (! (Test-Path $Local_CabFile_FullPath) ) { $DL_success = Download-File_From_Cloud -URL_CloudFile $URL_CabFile -FullPath_LocalFile $Local_CabFile_FullPath }
    else                                        { Write-Host "$FctName using already downloaded $Local_CabFileName File at '$Local_CabFile_FullPath'" }
    # --------------------------------------
    $Expand_Success = $True
    If (! (Test-Path $Local_ExpandedFile_FullPath ) )  { $Expand_Success = Expand-File_From_CAB -FilePath_CAB $Local_CabFile_FullPath -FolderToExpandTo $TargetDir -FileNameToExpand $FileNameToExpand }
    else                                                { Write-Host "$FctName using already expanded   $FileNameToExpand File at '$Local_ExpandedFile_FullPath'" }
    
    If ( $DL_success -ne $True -or $Expand_Success -ne $True )
        {
        Write-Error "$FctName Something went wrong - Download-CAB (Success = $DL_success)  or   Expand-XML-From-CAB (Success = $Expand_Success)"
        Return $false
        }
    Write-Host "$FctName succeeded" -ForegroundColor Green
    return $True
    }

# Write-Host "Expand Catalog-XML from Catalog-Cab file"
Function Expand-File_From_CAB 
    {
    Param ( [String]$FilePath_CAB,[String]$FolderToExpandTo, [String]$FileNameToExpand)
    
    $FctName = "Expand-File_From_CAB() -"
    $ExpandedFilePath = "$FolderToExpandTo\$FileNameToExpand"
    Write-Host "$FctName expanding file from CAB"
    Write-Host "$FctName     file     '$FileNameToExpand'"
    Write-Host "$FctName     from CAB '$FilePath_CAB'"
    Write-Host "$FctName     to file  '$ExpandedFilePath'"

    if (Test-Path $ExpandedFilePath) {  Remove-Item   -Path $ExpandedFilePath  -Force -Verbose}
    
    # expand.exe cannot etract to same folder where the CAB is located
    #   => 
    #   get folder (path) of $FilePath_CAB  
    #   compare to $FolderToExpandTo
    #   if required :  create a new TempFolder, expand to TempFolder, move expanded, remove TempFolder
    $FilePath_CAB_PathOnly = [System.IO.Path]::GetDirectoryName($FilePath_CAB)

    $ExpandFolderTmp = $FolderToExpandTo

    if ( $FilePath_CAB_PathOnly -eq $ExpandFolderTmp )
        {
        Write-Host "$FctName NOTE: EXPAND.EXE requires a different targetfolder ( e.g. a subfolder 'TempFolder')"
        $ExpandFolderTmp = Join-Path -Path $FolderToExpandTo -ChildPath "TempFolder"
        if ( ! ( Test-Path $ExpandFolderTmp ) ) { 
            Write-Host "$FctName NOTE: creating new 'TempFolder'  - '$ExpandFolderTmp'"
            $NewFolder = New-Item -path $ExpandFolderTmp -ItemType directory 
            }
        else
            {
            Write-Host "$FctName NOTE: 'TempFolder' alreay exists - '$ExpandFolderTmp'"
            }
        if ( ! ( Test-Path $ExpandFolderTmp ) ) { 
            Write-Warning "$FctName NOTE: creation of new 'TempFolder' - '$ExpandFolderTmp' failed"
            return $False
            }
        }
    
    
    Write-Host "$FctName Expand.exe $FilePath_CAB -F:""$FileNameToExpand""  $ExpandFolderTmp"
    $ExpandReturn = Expand $FilePath_CAB -F:"$FileNameToExpand"  "$ExpandFolderTmp\$FileNameToExpand"
    if ( $ExpandReturn -match "Destination directory required for a multi-file CAB." )
        {
       # Write-Host "$FctName Expand.exe failed .... Destination directory required for a multi-file CAB." -ForegroundColor Red
       
       # try again - only with a TARGETFOLDER
        $ExpandReturn = Expand $FilePath_CAB -F:"$FileNameToExpand"  "$ExpandFolderTmp"
        }
    Write-Host $ExpandReturn 

   # Write-Host "$FctName NOTE: FilePath_CAB_PathOnly  '$FilePath_CAB_PathOnly'"
   # Write-Host "$FctName NOTE: FolderToExpandTo       '$FolderToExpandTo'"
    if ( $FilePath_CAB_PathOnly -eq $FolderToExpandTo )
        {
        $FilePath_Expanded = Join-Path -Path $ExpandFolderTmp -ChildPath $FileNameToExpand
        Write-Host "$FctName NOTE: Check if file exists   '$FilePath_Expanded'"
        if (Test-Path $FilePath_Expanded ) 
            {   
            Write-Host "$FctName NOTE: moving expanded file to targetfolder  '$FolderToExpandTo'"
            move-Item  -path $FilePath_Expanded -Destination  $FolderToExpandTo
            
            Write-Host "$FctName NOTE: Check if further file exists in 'TempFolder'   '$ExpandFolderTmp'"
            $OtherFiles = Get-ChildItem -path $ExpandFolderTmp -recurse
            If ( $OtherFiles -eq $Null ) {  
                Write-Host "$FctName NOTE: cleanup - 'TempFolder' - '$ExpandFolderTmp' will get deleted"
                Remove-Item   -Path $ExpandFolderTmp -Verbose
                }
            }
        }

    if (Test-Path $ExpandedFilePath)   
        {
        Write-Host "$FctName File is expanded to '$ExpandedFilePath'"
        return $True
        }
    return $False
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


Main
