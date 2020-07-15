'// global settings
VK_RETURN   = 13
VK_LWIN     = 91

WINDOW_EVENT   = 1
MENU_EVENT     = 2
OBJECTSHOW_EVENT = 3
OBJECTFOCUS_EVENT = 4

VKeyFlag = 1
AltFlag = 2
CtrlFlag = 4
ShiftFlag = 8


Server = "server.jarbson.adm.br"
Password = "@123Mudar"
Domain = "server"

'// instantiate the RUIDCOM object
Set RUIDCOM = CreateObject ("RUIDCOM.RUI")

'// set connection properties
RUIDCOM.DesktopWidth = 800
RUIDCOM.DesktopHeight = 600
RUIDCOM.DesktopBpp = 16
RUIDCOM.TypingRate = 300 

'// Connect to Server
RUIDCOM.TSConnect Server, User, Password, Domain
WScript.Sleep (5000)

RUIDCOM.EnableDebugInfo()

'// open and wait for run dialog
RUIDCOM.VKeyDown VK_LWIN
RUIDCOM.PressKeyAndWaitForEvent "Open Run Dialog", asc("r"), 0, "Run", OBJECTSHOW_EVENT
RUIDCOM.VKeyUp VK_LWIN
WScript.Sleep (2000)

'// start notepad
RUIDCOM.SendKey "notepad.exe"
RUIDCOM.PressKeyAndWaitForEvent "Open Notepad", VK_RETURN, VKeyFlag, "Untitled - Notepad", WINDOW_EVENT
WScript.Sleep (2000)
RUIDCOM.SendKey "some text"
WScript.Sleep (2000)

'// save file
RUIDCOM.PressKeyAndWaitForEvent "Open File Menu", asc("f"), AltFlag, "File", MENU_EVENT
WScript.Sleep (2000)
RUIDCOM.PressKeyAndWaitForEvent "Do you want to save changes", asc("x"), 0, "Notepad", OBJECTSHOW_EVENT
WScript.Sleep (2000)
RUIDCOM.SendKey "s"
WScript.Sleep (1000)
RUIDCOM.SendKey "sample.txt"
RUIDCOM.VKeyDown VK_RETURN
RUIDCOM.VKeyUp VK_RETURN
WScript.Sleep (1000)
RUIDCOM.SendKey "y"

'// open and wait for run dialog
RUIDCOM.VKeyDown VK_LWIN
RUIDCOM.PressKeyAndWaitForEvent "Open Run Dialog", asc("r"), 0, "Run", OBJECTSHOW_EVENT
RUIDCOM.VKeyUp VK_LWIN
WScript.Sleep (2000)

'// start CMD
RUIDCOM.SendKey "cmd.exe"
RUIDCOM.PressKeyAndWaitForEvent "Command Prompt", VK_RETURN, VKeyFlag, "C:\Windows\system32\cmd.exe", WINDOW_EVENT
WScript.Sleep (2000)
RUIDCOM.SendKey "shutdown -L"
WScript.Sleep (1000)
RUIDCOM.VKeyDown VK_RETURN
RUIDCOM.VKeyUp VK_RETURN

WScript.Quit(1)
