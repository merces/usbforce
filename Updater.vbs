' USBForce (Updater) - Mente Binária - Fernando Mercês
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

If WScript.Arguments.Count = 1 Then
	If WScript.Arguments.Item(0) = "/s" Then
		Call checaUpdate(Versao, True)
	End If
Else
	Call checaUpdate(Versao, False)
End If

WScript.Quit


' Procedimento que checa por atualizações
Sub checaUpdate(Versao, Silencioso)

On Error Resume Next

Set objHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "GET", "http://www.mentebinaria.com.br/usbforce/versao.txt", False
objHTTP.Send

If Err.Number <> 0 Then
	If Not Silencioso Then
		MsgBox "Falha ao conectar-se ao servidor.", vbCritical, "USBForce " + Versao
	End If
	WScript.Quit
End If

If objHTTP.Status = 200 Then
s = objHTTP.ResponseText

	posTab = InStr(1, s, vbTab, vbTextCompare)
	
	If posTab > 0 Then
		versaoSite = Left(s, posTab - 1)
	End If
	
	s = Mid(s, posTab + 1)
	posTab = InStr(1, s, vbTab, vbTextCompare)
	
	If posTab > 0 Then
		dataSite = Mid(s, 1, posTab - 1)
	End If
	
End If

If versaoSite > Versao Then
	
	r = MsgBox("Nova versão do USBForce disponível." & vbCrLf & vbCrLf _
			& "Deseja atualizar para a versão " & versaoSite & " de "_
			& dataSite & " (você será redirecionado para a página"_
			& " de download)?", vbYesNo + vbQuestion, "USBForce " + Versao)

	If r = 6 Then
	
		Set objShell = CreateObject("Wscript.Shell")
		objShell.Run "http://www.mentebinaria.com.br/remository/func-startdown/3/"
		
	End If
	
Else
	
	If Not Silencioso Then
		MsgBox "Você possui a versão mais recente do USBForce.", vbInformation, "USBForce " + Versao
	End If
	
End If

End Sub