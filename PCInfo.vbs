' Used to expand environment variables
Set sh = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
Set Loc = CreateObject("WbemScripting.SWbemLocator")
Set Svc = Loc.ConnectServer(".", "root\cimv2")
Svc.Security_.ImpersonationLevel = 3

'Function for selecting folder

Dim strPath
strPath = SelectFolder( "" )
If strPath = vbNull Then
    WScript.Echo "Cancelled"
Else
    WScript.Echo "Selected Folder: """ & strPath & """"
End If

Function SelectFolder( myStartFolder )

    ' Standard housekeeping
    Dim objFolder, objItem, objShell
    
    ' Custom error handling
    On Error Resume Next
    SelectFolder = vbNull

    ' Create a dialog object
    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Select Folder", 0, myStartFolder )

    ' Return the path of the selected folder
    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

    ' Standard housekeeping
    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function

FileNameRaw = strPath & "\systeminfo.txt"

Set FileRaw = fs.CreateTextFile(FileNameRaw, true, true)

USRDomain = sh.ExpandEnvironmentStrings( "%USERDOMAIN%" )
USRName = sh.ExpandEnvironmentStrings("%USERNAME%")
strComputer = "." 

FileRaw.WriteLine "User: " & USRDomain & "\" & USRName & vbCrLf

FileRaw.WriteLine "System Information:" 
FileRaw.WriteLine "-------------------" 

' OS Caption
WScript.Echo "Getting OS Caption..."
Set OSCaps = Svc.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")
i = 0
For each OSCap in OSCaps
    i = i + 1
    FileRaw.WriteLine "Caption: " & OSCap.Caption 
Next
If i = 0 Then
    FileRaw.WriteLine "Caption is not found" 
End If
Set OSCaps = Nothing

' OS Version
WScript.Echo "Getting OS Version..."
Set OSVers = Svc.ExecQuery("SELECT Version FROM Win32_OperatingSystem")
i = 0
For each OSVer in OSVers
    i = i + 1
    FileRaw.WriteLine "Version: " & OSVer.Version 
Next
If i = 0 Then
    FileRaw.WriteLine "Version is not found" 
End If
Set OSVers = Nothing

' OS Architecture
WScript.Echo "Getting OS Architecture..."

Dim system_architecture
Dim process_architecture

Set shProcEnv = sh.Environment("Process")

process_architecture= shProcEnv("PROCESSOR_ARCHITECTURE") 

If process_architecture = "x86" Then    
    system_architecture= shProcEnv("PROCESSOR_ARCHITEW6432")

    If system_architecture = ""  Then    
        system_architecture = "x86"
    End if    
Else    
    system_architecture = process_architecture    
End If

FileRaw.WriteLine "OS Architecture: " & system_architecture

' Hostname
WScript.Echo "Getting Hostname..."
PCName = sh.ExpandEnvironmentStrings("%COMPUTERNAME%")
FileRaw.WriteLine "Hostname: " & PCName

' IP Addresses
WScript.Echo "Getting IP Address..."
strQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress > ''"

Set objWMIService = GetObject( "winmgmts://./root/CIMV2" )
Set colItems = objWMIService.ExecQuery( strQuery, "WQL", 48 )

For Each objItem In colItems
    If IsArray( objItem.IPAddress ) Then
        If UBound( objItem.IPAddress ) = 0 Then
            strIP = objItem.IPAddress(0)
        Else
            strIP = Join( objItem.IPAddress, "," )
        End If
    End If
Next

FileRaw.WriteLine "IP Address: " & strIP

' CPU name
WScript.Echo "Getting CPU Info..."
Set CPUs = Svc.ExecQuery("SELECT Name FROM Win32_Processor")
i = 0
For each CPU in CPUs
    i = i + 1
    FileRaw.WriteLine "Processor: " & CPU.Name
Next
If i = 0 Then
    FileRaw.WriteLine "No CPUs found"
End If
Set CPUs = Nothing

' RAM Size
WScript.Echo "Getting RAM Size..."
Set RAMs = Svc.ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
i = 0
GB = 1024 * 1024 * 1024
For each RAM in RAMs
    i = i + 1
    FileRaw.WriteLine "RAM (Gb): " & Round(RAM.TotalPhysicalMemory / GB,3)
Next
If i = 0 Then
    FileRaw.WriteLine "RAM size not found"
End If
Set RAMs = Nothing

' Free Physical Memory
WScript.Echo "Getting Free Physical Memory..."
Set FPMs = Svc.ExecQuery("Select * from Win32_OperatingSystem")
i = 0
MB = 1024 * 1024
For each FPM in FPMs
    i = i + 1
    FileRaw.WriteLine "Free Physical Memory (Mb): " & Round(FPM.FreePhysicalMemory / MB,3)
Next
If i = 0 Then
    FileRaw.WriteLine "Free Physical Memory is not found"
End If
Set FPMs = Nothing

' Free Virtual Memory
WScript.Echo "Getting Free Virtual Memory..."
Set FVMs = Svc.ExecQuery("Select * from Win32_OperatingSystem")
i = 0
MB = 1024 * 1024
For each FVM in FVMs
    i = i + 1
    FileRaw.WriteLine "Free Virtual Memory (Mb): " & Round(FVM.FreeVirtualMemory / MB,3) & vbCrLf
Next
If i = 0 Then
    FileRaw.WriteLine "Free Virtual Memory is not found" & vbCrLf
End If
Set FPMs = Nothing

' Disk info
WScript.Echo "Getting Disk Information..."
FileRaw.WriteLine "Disk Info:"
FileRaw.WriteLine "----------" 
Set colDiskDrives = objWMIService.ExecQuery ("Select * from Win32_DiskDrive")
i = 0
GB = 1024 * 1024 * 1024
For each objDiskDrive in colDiskDrives
    i = i + 1
    FileRaw.WriteLine "#" & i & " Caption: " & objDiskDrive.Caption & " | Name: " & objDiskDrive.Name & _
        " | Type: " & objDiskDrive.MediaType & " | Size (Gb): " & Round(objDiskDrive.Size / GB,3)
Next
FileRaw.WriteLine vbCrLf

' Outlook Info

WScript.Echo "Getting Outlook Information..."
FileRaw.WriteLine "Outlook Info:"
FileRaw.WriteLine "-------------" 

' For Outlook 2007, 2010

MySoftware = "Outlook"
Set colSoftware = objWMIService.ExecQuery ("Select * from Win32_Product where Name Like " & CommandLineLike(MySoftware))
i = 0
For Each objSoftware in colSoftware
    i = i + 1
    FileRaw.WriteLine "Caption: " & objSoftware.Caption
    FileRaw.WriteLine "Version: " & objSoftware.Version & vbCrLf
Next

' For Outlook 2003

If i = 0 Then
    MySoftware = "Office"
    Set colSoftware = objWMIService.ExecQuery ("Select * from Win32_Product where Name Like " & CommandLineLike(MySoftware))
    For Each objSoftware in colSoftware
        FileRaw.WriteLine "Caption: " & objSoftware.Caption
        FileRaw.WriteLine "Version: " & objSoftware.Version & vbCrLf
    Next
End If

Function CommandLineLike(MySoftware)   
    MySoftware = Replace(MySoftware, "\", "\\")   
    CommandLineLike = "'%" & MySoftware & "%'"   
End Function

' Kaspersky Info

WScript.Echo "Getting Kaspersky Information..."
FileRaw.WriteLine "Kaspersky Info:"
FileRaw.WriteLine "---------------" 

MySoftware = "Kaspersky"
Set colSoftware = objWMIService.ExecQuery ("Select * from Win32_Product where Name Like " & CommandLineLike(MySoftware))
i = 0
For Each objSoftware in colSoftware
    i = i + 1
    FileRaw.WriteLine "Caption: " & objSoftware.Caption
    FileRaw.WriteLine "Version: " & objSoftware.Version & vbCrLf
Next
If i = 0 Then 
     FileRaw.WriteLine "Kaspersky is not installed" & vbCrLf
End If

' Path for MAPI lib

WScript.Echo "Getting MAPI dll Information..."
FileRaw.WriteLine "MAPI dll Library Info:"
FileRaw.WriteLine "----------------------" 

Set mapiFile = fs.GetFile("C:\WINDOWS\system32\cgmxui32.dll")

FileRaw.WriteLine "Path for MAPI library: " & mapiFile.Path & vbCrLf

' Get eventviewer logs (today + 2 previous days)
WScript.Echo "Copying event viewer logs for last 2 days ..."

MyDate = DateAdd("d", -3, Now())

On Error Resume Next


Set objFileApp = fs.CreateTextFile(strPath & "\applog.csv", True)
Set objFileSys = fs.CreateTextFile(strPath & "\systemlog.csv", True)
Set objFileSec = fs.CreateTextFile(strPath & "\securitylog.csv", True)

ServerTime = Now    

' Collect logs in rows
intRecordNum = 0
row = 0
row1 = 0
row2 = 0

' WMI Core Section 
Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate,(Security)}!\\" & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMI.ExecQuery _ 
       ("SELECT * FROM Win32_NTLogEvent where Logfile = 'Application' and " _ 
        & "TimeWritten > '" & MyDate & "'")

Set colLoggedEvents1 = objWMI.ExecQuery _ 
       ("SELECT * FROM Win32_NTLogEvent where Logfile = 'System' and " _ 
        &  "TimeWritten > '" & MyDate & "'")
        
Set colLoggedEvents2 = objWMI.ExecQuery _ 
       ("SELECT * FROM Win32_NTLogEvent where Logfile = 'Security' and " _ 
        &  "TimeWritten > '" & MyDate & "'")


' Next section loops through ID properties

For Each objItem in colLoggedEvents                                                                                        
    objFileApp.WriteLine("Logfile: " & objItem.Logfile & "," & " source " & objItem.SourceName & "," & _
        "Message: " & objItem.Message & "," & _
        "TimeGenerated: " & WMIDateStringToDate(objItem.TimeGenerated) )
Next


For Each objItem1 in colLoggedEvents1
    objFileSys.WriteLine("Logfile: " & objItem1.Logfile & "," & " source " & objItem1.SourceName & "," & _
        "Message: " & objItem1.Message & "," & _
        "TimeGenerated: " & WMIDateStringToDate(objItem1.TimeGenerated) )
Next                      

For Each objItem1 in colLoggedEvents2
    objFileSec.WriteLine("Logfile: " & objItem1.Logfile & "," & " source " & objItem1.SourceName & "," & _
        "Message: " & objItem1.Message & "," & _
        "TimeGenerated: " & WMIDateStringToDate(objItem1.TimeGenerated) )
Next                                               

Function WMIDateStringToDate(dtmDate) 
 WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _ 
 Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _ 
 & " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2)) 
End Function

objFile.Close
Set objFile = Nothing


' Copy MAPI log file

WScript.Echo "Creating MAPI log file copy..."
FileRaw.WriteLine "MAPI log file:"
FileRaw.WriteLine "--------------"

FilePath = sh.ExpandEnvironmentStrings("%TEMP%") &"\cgmxp.log"

If (fs.FileExists(FilePath)) Then
    FileRaw.WriteLine FilePath & " exists."
    fs.CopyFile FilePath, strPath + "\", True
Else
    FileRaw.WriteLine "MAPI log file does not exist! Verify that logging at MAPI connector is enable."
End If

FileErr.Close
FileRaw.Close

WScript.Echo "Finished!"

Set app = Nothing

Set fs = Nothing
Set sh = Nothing