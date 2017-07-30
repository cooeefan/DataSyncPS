#Get Config file and initial parameters
[xml]$ConfigFile = Get-Content $PSScriptRoot"\config.xml"
$LogFilePath = $PSScriptRoot + $ConfigFile.Setting.LogPath + "\" + (Get-Date -UFormat "%Y%m%d").ToString() + ".txt"
$SQLUploadFileFolder = $PSScriptRoot + $ConfigFile.Setting.SQLUploadFileFolder
$ArchFolder = $PSScriptRoot + $ConfigFile.Setting.ArchFolder
$ToolsFolder = $PSScriptRoot + $ConfigFile.Setting.ToolsFolder
$TargetServer = $ConfigFile.Setting.TargetServer
$TargetDB = $ConfigFile.Setting.TargetDB


Function LogWriter($LogString)
{
    $LogString = (Get-Date -Format G).ToString() + " : "+ $LogString
    IF ($ConfigFile.setting.LogWriter -eq "ON")
    {
        $LogString | Out-File -FilePath $LogFilePath -Append
    }
    Write-Host $LogString
}

Function ExeSQL($SQLString)
{
    $connectionString = "Data Source=$TargetServer; " +"Integrated Security=SSPI; " +"Initial Catalog=$TargetDB"

    $Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = $connectionString
    $Connection.Open()

    $Datatable = New-Object System.Data.DataTable
    $Command = New-Object System.Data.SQLClient.SQLCommand

    $Datatable = New-Object System.Data.DataTable
    $Command = New-Object System.Data.SQLClient.SQLCommand
    Try
    {
        $Command.Connection = $Connection
        $Command.CommandText = $SQLString
        $Command.CommandTimeout = 0
        $Reader=$Command.ExecuteReader()
        $Datatable.Load($Reader)
        LogWriter(-join($SQLString," execute sucessfully"))
        Return $Datatable
    }
    Catch
    {
        $ErrorMessage = $_.Exception.Message
        LogWriter(-join($SQLString," DIDN'T execute sucessfully. Below is the error message"))
        LogWriter($ErrorMessage)
    }
    $Connection.Close()
}

Function UnzipFile($FilePath)
{
    $ZipExec = -join($ToolsFolder,"\7za.exe")
    $ZipArgument = -join("e ",$FilePath.ToString()," -o",$SQLUploadFileFolder,"\"," -aoa")
	$TempFile = [System.IO.Path]::GetTempFileName()

	$Process = (Start-Process -FilePath $ZipExec -ArgumentList $ZipArgument -Wait -PassThru -RedirectStandardOutput $TempFile)

	If($Process.ExitCode -ne 0)
	{
		LogWriter(-join("File ",$FilePath," didn't unzip sucessfully, below is the error message"))
		$ErrorText = Get-Content $TempFile
		LogWriter($ErrorText)
        Return $false
	}

	Else
	{
		LogWriter(-join("File ",$FileList[$i]," unzipped sucessfully"))
        Return $true
	}

	Remove-Item $TempFile

}

Function UnixFileCheck($FilePath)
{
    $reader = new-object System.io.streamreader($FilePath)
    $LineText = @()


    while(($reader.Peek()) -gt 0)
    {
        $LineText = $LineText + $reader.Read()
        IF(($LineText[-1] -eq 10) -and ($LineText[-2] -ne 13))
        {
            $reader.Close()
            Return $True
            Break
        }
        IF(($LineText[-1] -eq 10) -and ($LineText[-2] -eq 13))
        {
            $reader.Close()
            Return $False
            Break
        }
    }
    Return $False
    $reader.Close()
}

Function ConvertUNIXFile($FilePath)
{
    $FileConvertExec = -join($ToolsFolder,"\unix2dos.exe")
	$TempFile = [System.IO.Path]::GetTempFileName()

    $Process = (Start-Process -FilePath $FileConvertExec -ArgumentList $FilePath -Wait -PassThru -RedirectStandardError $TempFile)

    If($Process.ExitCode -ne 0)
	{
		LogWriter(-join("File ",$FilePath," didn't convert sucessfully, below is the error message"))
		$ErrorText = Get-Content $TempFile
		LogWriter($ErrorText)
        Return $false
	}

	Else
	{
		LogWriter(-join("File ",$FilePath," convert sucessfully"))
        Return $true
	}

	Remove-Item $TempFile
}

Function UploadFile($FilePath)
{
    $FileName = $FilePath.Substring($FilePath.LastIndexOf("\")+1,($FilePath.LastIndexOf(".")-$FilePath.LastIndexOf("\")-1))
    LogWriter(-join("File ",$FileName," start uploading..."))

    #Get Table Name
    IF($FileName.indexOf("-UAT") -gt 0)
    {
        $TableName = $FileName.Substring(0,$FileName.LastIndexOf("-UAT"))
    }
    Else
    {
        $TableName = $FileName
    }


    #Create table to upload
    $FileContent = Get-Content -Path $FilePath -TotalCount 10

    $EmptyStringIndex = 0
    $VerticalBarIndex = 0

    For ($i=0; $i -lt 11; $i++)
    {
        IF($FileContent[$i].Indexof("|") -eq -1)
        {
            $EmptyStringIndex = $i            
        }
        IF($FileContent[$i].Indexof("|") -gt -1)
        {
            $VerticalBarIndex = $i
            break            
        }

    }
    
    #Check File is from an empty table
    IF($FileContent[0] -and [string]::IsNullOrEmpty($FileContent[1]) -and ($VerticalBarIndex -eq 0))
    {
        LogWriter(-join($FilePath," is a file from empty table or view"))
        $EmptyFileContent = Get-Content -Path $FilePath -TotalCount 1
        $SQLString = ""
        $SQLString = -join (" IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[",$TableName,"]') AND type in (N'U'))DROP TABLE [dbo].[",$TableName,"] Create Table [dbo].[",$TableName,"]( [",$EmptyFileContent.Replace("|","] VARCHAR(250), ["),"] VARCHAR(250))")
        #LogWriter($SQLString)
        ExeSQL($SQLString)

    }

    #Check file is from table or view directly
    ELSEIF($FileContent[0] -and $FileContent[1] -and ($FileContent[0].IndexOf("|") -gt -1) -and ($FileContent[1].IndexOf("|") -gt -1))
    {
        LogWriter(-join($FilePath," is a regular file from table or view"))
        $SQLString = ""
        $SQLString = -join (" IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[",$TableName,"]') AND type in (N'U'))DROP TABLE [dbo].[",$TableName,"] Create Table [dbo].[",$TableName,"]( [",$FileContent[0].Replace("|","] VARCHAR(250), ["),"] VARCHAR(250))")
        #LogWriter($SQLString)
        ExeSQL($SQLString)

    }

    #Check file is from table or view directly but only has one column
    ELSEIF($FileContent[0] -and $FileContent[1] -and ($VerticalBarIndex -eq 0))
    {
        LogWriter(-join($FilePath," is from table or view directly but only has one column"))
        $SQLString = ""
        $SQLString = -join (" IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[",$TableName,"]') AND type in (N'U'))DROP TABLE [dbo].[",$TableName,"] Create Table [dbo].[",$TableName,"]( [",$FileContent[0],"] VARCHAR(250))")
        #LogWriter($SQLString)
        ExeSQL($SQLString)
    }

    #Check file is from Calc
    ELSEIF(($EmptyStringIndex -gt 0) -and ($VerticalBarIndex -gt 0) -and ($VerticalBarIndex = $EmptyStringIndex+1) )
    {
        LogWriter(-join($FilePath," is from Calc"))
        $SQLString = ""
        $SQLString = -join (" IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[",$TableName,"]') AND type in (N'U'))DROP TABLE [dbo].[",$TableName,"] Create Table [dbo].[",$TableName,"]( [",$FileContent[$VerticalBarIndex].Replace("|","] VARCHAR(250), ["),"] VARCHAR(250))")
        #LogWriter($SQLString)
        ExeSQL($SQLString)
    }

    ELSE
    {
        LogWriter("File is empty or format is wrong, skip the file")
        Return $false
    }

    $BCPString = ""
    $BCPString = -join($TargetDB,".dbo.[",$TableName,"] in ",$FilePath," -F 2 -c -t ","""|"" -r\n -T -S ",$TargetServer)
    LogWriter($BCPString)
    $TempFile = [System.IO.Path]::GetTempFileName()
    $Process = (Start-Process BCP.exe -ArgumentList $BCPString -Wait -PassThru -RedirectStandardError $TempFile)

    If($Process.ExitCode -ne 0)
	{
		LogWriter(-join("File ",$FilePath," didn't BCP sucessfully, below is the error message"))
		$ErrorText = Get-Content $TempFile
		LogWriter($ErrorText)
	    Remove-Item $TempFile
        Return $false
	}

	Else
	{
		LogWriter(-join("File ",$FilePath," BCP sucessfully"))
	    Remove-Item $TempFile
        Return $true
	}





}

############################ MAIN PROCESS ##############################################

LogWriter("")
LogWriter("")
LogWriter("Process Start...............................")
LogWriter("LogFilePath: "+$LogFilePath)
LogWriter("LogFile is: "+$ConfigFile.setting.LogWriter)
LogWriter("SQLUploadFileFolder: "+$SQLUploadFileFolder)
LogWriter("ArchFolder: "+$ArchFolder)
LogWriter("TargetServer: "+$TargetServer)
LogWriter("TargetDB: "+$TargetDB)


#Get File list which needs to process
$FileList = @(Get-ChildItem -Path $SQLUploadFileFolder -Name -file)
For ($i=0; $i -lt $FileList.Length; $i++)
{
    #Write-Host (-join($SQLUploadFileFolder,"\",$FileList[$i]))
    $FileList[$i] = -join($SQLUploadFileFolder,"\",$FileList[$i])
}
LogWriter("Files need to Process: " +$FileList)

#Start Processing Files

For ($i=0; $i -lt $FileList.Length; $i++)
{
    LogWriter(-join("Start Processing File: ",$FileList[$i]))

    #Process .gz file
    IF($FileList[$i].ToString().Substring($FileList[$i].ToString().LastIndexOf(".")+1) -eq "gz")
	{
        LogWriter(-join($FileList[$i]," is a zipped file, start unzipping"))
        IF(UnzipFile($FileList[$i]))
        {
            IF(UnixFileCheck($FileList[$i].ToString().Replace(".gz","")))
                {
                    LogWriter(-join($FileList[$i].ToString().Replace(".gz",""))," is a Unix file. Start converting...")
                    IF(ConvertUNIXFile($FileList[$i].ToString().Replace(".gz","")))
                    {
                        LogWriter(-join("Start Uploading File: ",$FileList[$i].ToString().Replace(".gz","")))
                        UploadFile($FileList[$i].ToString().Replace(".gz",""))
                        Remove-Item ($FileList[$i].ToString().Replace(".gz",""))
                    }
                    Else
                    {
                        Continue
                    }

                }
            Else
            {
                UploadFile($FileList[$i].ToString().Replace(".gz",""))
                Remove-Item ($FileList[$i].ToString().Replace(".gz",""))
            }
        }
        Else
        {
            LogWriter(-join("File: ",$FileList[$i]," Didn't unzip sucessfully, skip the file"))
            Continue
        }

	}

    #Process .csv file
    ELSEIF($FileList[$i].ToString().Substring($FileList[$i].ToString().LastIndexOf(".")+1) -eq "csv")
    {
        LogWriter(-join($FileList[$i]," is a .csv file, start converting"))
        $tempCSVTxtFile = $FileList[$i].ToString().Replace('.csv','.txt')
        Import-Csv $FileList[$i] | Export-Csv -path $tempCSVTxtFile -Delimiter "|"
        #Import-Csv 'C:\Users\fjiang1\OneDrive\Codes\PowerShell\NewDataSync\Download\cfgICSSpiff.csv' | Export-Csv -path 'C:\Users\fjiang1\OneDrive\Codes\PowerShell\NewDataSync\Download\cfgICSSpiff.txt' -Delimiter "|"
        (Get-Content -Path $tempCSVTxtFile).Replace('"','') | select -Skip 1 | Set-Content '$File-Temp' 
        Move '$File-Temp' $tempCSVTxtFile -Force
        LogWriter(-join($FileList[$i]," converted to .txt file"))

        LogWriter(-join("Start Uploading File: ",$tempCSVTxtFile))
        UploadFile($tempCSVTxtFile)
        
    } 

    #Process .txt file
    ELSEIF($FileList[$i].ToString().Substring($FileList[$i].ToString().LastIndexOf(".")+1) -eq "txt")
    {
        LogWriter(-join($FileList[$i]," is a .txt file, start Processing"))
        IF(UnixFileCheck($FileList[$i]))
            {
                LogWriter(-join($FileList[$i]," is a Unix file. Start converting..."))
                IF(ConvertUNIXFile($FileList[$i]))
                {
                    LogWriter(-join("Start Uploading File: ",$FileList[$i]))
                    UploadFile($FileList[$i])
                }
                Else
                {
                    Continue
                }
        
            }
        Else
        {
            LogWriter(-join("Start Uploading File: ",$FileList[$i]))
            UploadFile($FileList[$i])
        }
    } 

    #File CANNOT be processed
    ELSE
    {
        LogWriter(-join($FileList[$i]," is not the File can be processed by this programm, skip the file"))
        Continue
    }

    #Move the fils to Arch

}

#Move the file to Arch

LogWriter("Start archiving Uploaded files")
$MoveFiles = -join($SQLUploadFileFolder,'\*.*')

Move-Item -Path $MoveFiles -Destination $ArchFolder


LogWriter("Upload process finished")