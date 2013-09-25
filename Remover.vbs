' USBForce (Remover) - Mente Binária - Fernando Mercês
' Data de início do código: 26/1/2010 00:23
' Última alteração: 31/1/2010 20:38

' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.

Versao = "1.2"

Const HKLM = &H80000002

Sub removeProtecao()

	Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	
	Chave = "SOFTWARE\MenteBinaria\USBForce"
	objReg.DeleteKey HKLM, Chave
	
	'Desabilita o autorun para todas as unidades
	Chave = "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
	Valor = "NoDriveTypeAutoRun"
	objReg.DeleteValue HKLM, Chave, Valor
	
	'Desabilita execução do autorun.inf
	Chave = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\IniFileMapping\Autorun.inf"
	objReg.DeleteKey HKLM, Chave
	
	Set objRede = CreateObject("WScript.Network")
	escreveLog "1", "Proteção desabilitada pelo usuário " + objRede.UserName + "."

End Sub

Sub escreveLog(tipo, desc)

	Const LEITURA = 1
	Const ESCRITA = 2

	' Pega a data e hora atual no formato dd/mm/yyyy HH:MM
	Agora = FormatDateTime(Date, 2) & " " & Time
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject("Wscript.Shell")

	' Arquivo de log no diretório do programa
	logFile = "USBForce_Log.txt"
	
	' Se o arquivo de log não existir, é criado
	If Not objFSO.FileExists(logFile) Then
		
		objFSO.CreateTextFile(logFile)
		Set objLog = objFSO.OpenTextFile(logFile, ESCRITA)
		objLog.Write Agora & vbTab & "1" & vbTab & "Criado novo arquivo de log."
		objLog.Close
		
	Else
	
	' Adiciona uma entrada no início do arquivo de log
	Set objLog = objFSO.OpenTextFile(logFile, LEITURA)
	strContents = objLog.ReadAll
	objLog.Close
	Set objLog = objFSO.OpenTextFile(logFile, ESCRITA)
	objLog.WriteLine objLog.Write(Agora & vbTab & tipo & vbTab & desc)
	objLog.Write strContents
	objLog.Close
	
	End If

End Sub

' Desabilitar proteção
r = MsgBox("Tem certeza que deseja desbilitar a proteção do USBForce?",_
			vbYesNo + vbQuestion, "USBForce " +  Versao)
	    				
If r = 6 Then

   	Call removeProtecao
   	MsgBox "Proteção desativada. Faça logoff ou reinicie o PC."_
			, vbOKOnly + vbInformation, "USBForce " +  Versao
End If
	    
WScript.Quit