    
    #--------------------Part 1 - export autocomplete cache
    $userdirs = Get-ChildItem -path "c:\users\" -directory 
    foreach ($userdir in $userdirs) {
        $folder = "c:\users\$userdir\appdata\local\microsoft\outlook\roamcache"
        #see if folder exists
        if (Test-Path -Path $folder) {
            #allow access to folder
            $acl = get-acl -path $folder
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("sscv.local\codeblue","FullControl","Allow")
            $acl.SetAccessRule($Accessrule)
            $acl | set-acl "c:\users\$userdir\appdata\local\microsoft\outlook\roamcache"
            #see if there is an autocomplete file
            if (Test-Path -path $(Join-Path $folder "stream_autocomplete*")) {
                #trim .domain suffix in user folder name
                $cleanname = $($userdir.Name).split('.')[0]
                #create subfolder
                $newfolder = new-item "c:\autocomplete\$cleanname" -ItemType Directory
                #copy file unless preexisting
                copy-item (Join-Path $folder "stream_autocomplete*") -Destination $newfolder -exclude (Get-ChildItem $newfolder)
            }
        }
    }

    #--------------------Part 2 - import nk2
    If (-not (test-path -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover")) {
        New-item -path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -name "Autodiscover"
    }
    Set-ItemProperty -path "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover" -name "ZeroConfigExchange" -value 1
    Set-ItemProperty -path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -name "DefaultProfile" -value "Office 365"
    
    If (-not (test-path -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles")) {
        New-item -path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -name "Profiles"
    }
    
    If (-not (test-path -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Office 365")){
        New-item -path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles" -name "Office 365"
    }
    
    #find largest autocomplete cache backup
    $backupcache = get-childitem "c:\autocomplete\$env:username" -recurse -filter *stream_auto* | sort -descending -property length | select -first 1
    $backupcachedirectory = $backupcache.DirectoryName
    $nk2directory = "$env:APPDATA\Microsoft\Outlook"
    #if it doesn't exist already
    if (-not(Test-Path -path "$backupcachedirectory\Office 365.nk2" -pathtype leaf)) {
        #import to Outlook profile
        c:\autocomplete\nk2edit.exe /nk2_to_text $backupcache.fullname "$backupcachedirectory\Office 365.nk2"
        c:\autocomplete\nk2edit.exe /text_to_nk2 "$backupcachedirectory\Office 365.nk2" "$nk2directory\Office 365.nk2"

        #start outlook and close it, allowing profile to autoconfigure
        $outlook = start-process -filepath "C:\Program Files (x86)\Microsoft Office\root\Office16\outlook.exe" -ArgumentList '/importnk2','/profile "Office 365"' -passthru
        #wait 20 seconds
        start-sleep 20
        $outlook.kill()
    }






   
#-----------------------------   



#enable modern authentication
Set-ItemProperty -path "HKCU:\Software\Microsoft\Office\16.0\Common\Identity" -name "EnableADAL" -Value 1
#remove registry entry preventing easy autodiscover to office 365
Set-ItemProperty -path "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover" -name "ExcludeExplicitO365Endpoint" -Value 0
#enable zero config exchange (ZCE)