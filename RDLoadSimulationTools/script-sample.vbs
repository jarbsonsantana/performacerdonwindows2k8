'***************************************
'TS Scalability scenario script header
'***************************************

VK_RETURN   = 13
VK_MENU     = 18
VK_CONTROL  = 162
VK_TAB      = 9
VK_UP       = 38
VK_DOWN     = 40
VK_RIGHT    = 39
VK_LEFT     = 37
VK_SHIFT    = 16
VK_BACK     = 8
VK_ESCAPE   = 27
VK_SPACE    = 32
VK_END      = 35
VK_HOME     = 36
VK_DELETE   = 46
VK_LWIN     = 91

WINDOW_EVENT   = 1
MENU_EVENT     = 2
OBJECTSHOW_EVENT = 3
OBJECTFOCUS_EVENT = 4

VKeyFlag = 1
AltFlag = 2
CtrlFlag = 4
ShiftFlag = 8

sub ShowUsage

   WScript.Echo ("RUIDCOM ActiveX Script")
   WScript.Echo ("Usage:" + Chr(10) + Chr(10) + "<script> [- options]" + Chr(10))
   WScript.Echo ("Optional Parameters:"+ Chr(10) + Chr(10))
   WScript.Echo (" -s:server - The target TS server to use.")
   WScript.Echo (" -u:username - The username to use.")
   WScript.Echo (" -p:password - The default password to use.")
   WScript.Echo (" -d:domain - The default domain to use.")
   WScript.Echo (" -e:Exchange Server - The name of Exchange Server to be used.")
   WScript.Echo (" -a:<args> - Argument string which can be used in scripts.")
   WScript.Echo (" -f:# - Speed Factor. Determines how fast the script will run")
   WScript.Echo ("         Setting it to 1 sets typing speed to 35 WPM")
   WScript.Echo (" -x:# - Desktop Index" + vbLf)
   WScript.Quit(1)

end sub


SpecialDesktop = false

'***************************
'Parse the script arguments
'***************************

if WScript.Arguments.Count < 4  then
	WScript.Echo( "Invalid number of arguments!!" )
   ShowUsage
end if

ArgumentCount = WScript.Arguments.Count
dim argString


for indx = 0 to ArgumentCount - 1 step 1
   if InstrRev(WScript.Arguments(indx),"-s:") = 1 then
      argString = WScript.Arguments(indx)
      Server = Mid(argString,4)
   elseif InstrRev(WScript.Arguments(indx),"-u:") = 1 then
      argString = WScript.Arguments(indx)
      User = Mid(argString,4)
   elseif InstrRev(WScript.Arguments(indx),"-p:") = 1 then
      argString =  WScript.Arguments(indx)
      Password = Mid(argString,4)
   elseif InstrRev(WScript.Arguments(indx),"-d:") = 1 then
      argString =  WScript.Arguments(indx)
      Domain = Mid(argString,4)
   elseif InstrRev(WScript.Arguments(indx),"-e:") = 1 then
      argString =  WScript.Arguments(indx)
      ExchangeServer = Mid(argString,4)
   elseif InstrRev(WScript.Arguments(indx),"-f:") = 1 then
      argString =  WScript.Arguments(indx)
      numString = Mid(argString,4)
      SpeedFactor = CInt(numString)
   elseif InstrRev(WScript.Arguments(indx),"-x:") = 1 then
      argString =  WScript.Arguments(indx)
      DesktopID = Mid(argString,4)
      SpecialDesktop = true
   elseif InstrRev(WScript.Arguments(indx),"-a") = 1 then
      argString =  WScript.Arguments(indx)
      SpecialArg = Mid(argString,4)
   else
      WScript.Echo ("argument " + cstr(indx+1) + " is incorrect")
      WScript.Echo (WScript.Arguments(indx))
      ShowUsage
   end if
   
next

if SpecialDesktop = true Then
	RUIObject = "RUIDCOM.RUI"  + DesktopID
Else
	RUIObject = "RUIDCOM.RUI" 
End If

Err.Clear
Set RUIDCOM = CreateObject (RUIObject)
If Err.Number Then
    WScript.Echo("Error (" + _
                 Err.Number + _
                 "): unable to create RUIDCOM object!")
    WScript.Quit(1)
End If

RUIDCOM.TypingRate = 35*SpeedFactor


'// set connection properties
RUIDCOM.DesktopWidth = 800
RUIDCOM.DesktopHeight = 600
RUIDCOM.DesktopBpp = 16
RUIDCOM.TypingRate = 300 

'// Connect to Server
RUIDCOM.TSConnect Server, User, Password, Domain
WScript.Sleep (5000)

'// open and wait for run dialog
RUIDCOM.VKeyDown VK_LWIN
RUIDCOM.PressKeyAndWaitForEvent "Open Run Dialog", asc("r"), 0, "Run", OBJECTSHOW_EVENT
RUIDCOM.VKeyUp VK_LWIN
WScript.Sleep (2000)

'// start notepad
RUIDCOM.SendKey("notepad.exe")
RUIDCOM.PressKeyAndWaitForEvent "Open Notepad", VK_RETURN, VKeyFlag, "Untitled - Notepad", WINDOW_EVENT
WScript.Sleep (2000)
RUIDCOM.SendKey "some text"
WScript.Sleep (2000)

'// save file
RUIDCOM.PressKeyAndWaitForEvent "Open File Menu", asc("f"), AltFlag, "File", MENU_EVENT
WScript.Sleep (2000)
RUIDCOM.PressKeyAndWaitForEvent "Open Save As Dialog", asc("s"), 0, "Save as", OBJECTSHOW_EVENT
WScript.Sleep (2000)
RUIDCOM.SendKey "sample.txt"
RUIDCOM.PressKeyAndWaitForEvent "Confirm Save as", VK_RETURN, VKeyFlag, "Confirm Save", OBJECTSHOW_EVENT
WScript.Sleep (2000)
RUIDCOM.SendKey "y"

do while true 
	WScript.Sleep (2000)
Loop
