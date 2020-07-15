'// global settings
VK_RETURN   = 13
VK_LWIN     = 91
VK_TAB      = 9
VK_ESCAPE   = 27

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

'// open and wait for run dialog
RUIDCOM.VKeyDown VK_LWIN
RUIDCOM.PressKeyAndWaitForEvent "Open Run Dialog", asc("r"), 0, "Run", OBJECTSHOW_EVENT
RUIDCOM.VKeyUp VK_LWIN
WScript.Sleep (2000)

'// Abrindo sistema
RUIDCOM.SendKey "C:\PUBLICSOFT\Contabilidade\PS_Contabilidade.exe"
RUIDCOM.PressKeyAndWaitForEvent "Abrindo Contabilidade PublicSoft", VK_RETURN, VKeyFlag, "[PMAB]PublicSoft Contabilidade", WINDOW_EVENT
WScript.Sleep (15000)

'// Login
RUIDCOM.VKeyDown VK_TAB
RUIDCOM.VKeyUp VK_TAB
WScript.Sleep (500)
RUIDCOM.VKeyDown VK_TAB
RUIDCOM.VKeyUp VK_TAB
WScript.Sleep (500)
RUIDCOM.VKeyDown VK_TAB
RUIDCOM.VKeyUp VK_TAB
WScript.Sleep (500)
RUIDCOM.SendKey "senha"
WScript.Sleep (1000)
RUIDCOM.VKeyDown VK_RETURN
RUIDCOM.VKeyUp VK_RETURN
WScript.Sleep (30000)

do while true 

'// Executando rotinas básicas
'// Menu Arquivos>Orçamentário>QDD
RUIDCOM.PressKeyAndWaitForEvent "Menu Arquivo", asc("a"), AltFlag, "Application", MENU_EVENT
WScript.Sleep (3000)
RUIDCOM.SendKey "r"
WScript.Sleep (2000)
RUIDCOM.SendKey "q"
WScript.Sleep (4000)
RUIDCOM.VKeyDown VK_ESCAPE
RUIDCOM.VKeyUp VK_ESCAPE
WScript.Sleep (2000)
'// Menu Arquivos>Orçamentário>QDR
RUIDCOM.PressKeyAndWaitForEvent "Menu Arquivo", asc("a"), AltFlag, "Application", MENU_EVENT
WScript.Sleep (3000)
RUIDCOM.SendKey "r"
WScript.Sleep (2000)
RUIDCOM.SendKey "d"
WScript.Sleep (4000)
RUIDCOM.VKeyDown VK_ESCAPE
RUIDCOM.VKeyUp VK_ESCAPE
WScript.Sleep (2000)
'// Menu Relatorios>Balancetes Mensais
RUIDCOM.PressKeyAndWaitForEvent "Menu Relatorios", asc("r"), AltFlag, "Application", MENU_EVENT
WScript.Sleep (3000)
RUIDCOM.SendKey "b"
WScript.Sleep (2000)
RUIDCOM.VKeyDown VK_ESCAPE
RUIDCOM.VKeyUp VK_ESCAPE
WScript.Sleep (2000)
RUIDCOM.VKeyDown VK_ESCAPE
RUIDCOM.VKeyUp VK_ESCAPE
WScript.Sleep (2000)
'// Menu Relatorios>Balancetes Anuais
RUIDCOM.PressKeyAndWaitForEvent "Menu Relatorios", asc("r"), AltFlag, "Application", MENU_EVENT
WScript.Sleep (3000)
RUIDCOM.SendKey "l"
WScript.Sleep (2000)
RUIDCOM.VKeyDown VK_ESCAPE
RUIDCOM.VKeyUp VK_ESCAPE
WScript.Sleep (2000)
RUIDCOM.VKeyDown VK_ESCAPE
RUIDCOM.VKeyUp VK_ESCAPE
WScript.Sleep (2000)
'// Menu Ferramentas
RUIDCOM.PressKeyAndWaitForEvent "Menu Ferramentas", asc("f"), AltFlag, "Application", MENU_EVENT
WScript.Sleep (3000)
RUIDCOM.SendKey "g"
WScript.Sleep (12000)
RUIDCOM.VKeyDown VK_ESCAPE
RUIDCOM.VKeyUp VK_ESCAPE
WScript.Sleep (2000)
RUIDCOM.VKeyDown VK_ESCAPE
RUIDCOM.VKeyUp VK_ESCAPE
WScript.Sleep (2000)

Loop