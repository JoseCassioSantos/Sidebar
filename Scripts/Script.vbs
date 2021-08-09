



Sub Window_OnLoad()
window.resizeTo 220, 480

End Sub


Function BtnMin()
window.resizeTo 220, 60
End Function


Function BtnStart()

Set Apl =CreateObject("wscript.shell")


Apl.run("Taskmgr")

End Function


Function BtnEnd()

Set Apl =CreateObject("wscript.shell")


Apl.run("cmd")

End Function

Function Btnabrir()

Set Apl =CreateObject("wscript.shell")


Apl.run("explorer.exe")

End Function

Function Btncontrol()

Set Apl =CreateObject("wscript.shell")

Apl.run("control")

End Function


Function BtnSFTP()

Set Apl =CreateObject("wscript.shell")

Apl.run("\\10.221.243.12\ti_jbt\Scan\TI\SFTP\WinSCP.exe")

End Function

Function BtnVNC()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\ultravnc_1214\64\vncviewer.exe")

End Function


Function BtnCentreon()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\Links\Centreon.url")

End Function


Function BtnOcs()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\Links\OCS.url")

End Function



Function BtnNeoservice()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\Links\NeoService.url")

End Function



Function BtnNagios()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\Links\Nagios.url")

End Function

Function BtnGep()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\Links\Gep.Url")

End Function

Function BtnCortina()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\Links\Cortina.url")

End Function


Function BtnIntranet()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\Links\Intranet.Url")

End Function

Function BtnInside()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\Links\InsideTools.Url")
Apl.run("..\Links\InsideRelatorios.Url")

End Function

Function BtnCalc()
    Set Apl =CreateObject("wscript.shell")
    Apl.run("Calc")
End Function

Function BtnNotPad()
    Set Apl =CreateObject("wscript.shell")
    Apl.run("notepad")
End Function

Function BtnInfo()

Set Apl =CreateObject("wscript.shell")

Apl.run("msinfo32")

End Function

Function BtnMstsc()

Set Apl =CreateObject("wscript.shell")

Apl.run("mstsc")

End Function

Function BtnComplemento()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\complemento.hta")

End Function

Function BtnKasper()

set CMDuse = CreateObject("wscript.shell")

CMDuse.Run "cmd /c start ..\KperC\sky.bat",0, true

End Function


Function BtnNortel()

set CMDuse = CreateObject("wscript.shell")

CMDuse.Run "cmd /c call ..\Nortel\JDM3\JDM.BAT",0, true

End Function

Function BtnPaint()

Set Apl =CreateObject("wscript.shell")

Apl.run("mspaint")

End Function

Function BtnCMS()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\Cms\Rest_CMS.exe")

End Function

Function BtnCTI()

Set Apl =CreateObject("wscript.shell")

Apl.run("..\CTI\ServidorCTI.exe")

End Function




Function BtnMoni()

Set Apl =CreateObject("wscript.shell")


Apl.run("C:\Windows\system32\fsmgmt.msc")

End Function



Function BtnVer()

Set Apl =CreateObject("wscript.shell")


Apl.run("winver")

End Function


Function BtnPadrao()

Set Apl =CreateObject("wscript.shell")


Apl.run("ComputerDefaults")

End Function


Function Btnuser()

Set Rede =CreateObject("Wscript.Network")
Info = "Meu Dominio" & vbTab & "-> " & Rede.UserDomain & _ 
vbCrlf & "Meu Computador" & vbTab & "-> " & Rede.ComputerName & _
vbCrlf & "Meu Usuario" & vbTab & "-> " & Rede.UserName & vbCrlf
MsgBox(Info),32, "Informacoes de Rede"

End Function

Function BtnSobre()

Set help =CreateObject("Wscript.shell")
MsgBox "Ti NeoBpo JBT"& vbCrlf & _ 
	"Somos o setor de Tecnologia e informacoes local" & vbCrlf & _
	"Era pra ser Tecnologia DA informacao mais ninguem" & vbCrlf &_
	"Liga pra o Da nao e mesmo!!! ",32, "Sobre"

End Function

Function BtnIP()

set CMDuse = CreateObject("wscript.shell")

CMDuse.Run "cmd /c color f0 && ipconfig /all && pause && gpupdate"

End Function

Function BtnBat1()

set CMDuse = CreateObject("wscript.shell")

CMDuse.Run "cmd /c start \\10.221.243.13\",0, true

End Function

Function BtnBat2()

set CMDuse = CreateObject("wscript.shell")

CMDuse.Run "cmd /c start \\10.220.9.250\scan\",0, true

End Function

Function BtnBat3()

set CMDuse = CreateObject("wscript.shell")

CMDuse.Run "cmd /c start \\10.220.9.250\installer$",0, true

End Function



Function BtnSair()

window.close

End Function
