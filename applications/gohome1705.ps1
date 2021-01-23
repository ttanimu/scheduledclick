# scheduledclick / gohome1705.ps1 (based on clickat.ps1)
# https://github.com/ttanimu/scheduledclick
# 
# Steps
#   1.Click default button of a certain window if it exists.
#   2.Click second button of another certain window and retry twice.
#   3.Reboot PC.
#
# Parameters
#   TIME TO CLICK : 1705
#   1st window
#     WINDOW TITLE : It contains "SCE-IT-SVR ". 
#   2st window
#     WINDOW TITLE : It contains "TT10-P1100-10 ". 
#     PROCESS NAME : TT~P110010.V001.20170130.0000

$source = @"
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows

Public Class API_EnumWindows

	Public count As Integer

	Public Delegate Function EnumWindowsCallbackNormal( _
	ByVal hWnd As Integer, _
	ByVal lParam As Integer) As Boolean

	<DllImport("user32.dll", CharSet:=CharSet.Auto)> _
	 Public Shared Function EnumWindows( _
	 ByVal callback As EnumWindowsCallbackNormal, _
	 ByVal lParam As Integer) As Integer
	End Function

	Public Delegate Function EnumWindowsCallbackString( _
	   ByVal hWnd As Integer, _
	   ByVal lParam As String) As Boolean

	<DllImport("user32.dll", EntryPoint:="EnumWindows", CharSet:=CharSet.Ansi)> _
	   Public Shared Function EnumWindowsString( _
	   ByVal callback As EnumWindowsCallbackString, _
	   ByVal lParam As String) As Integer
	End Function

	<DllImport("user32.dll", CharSet:=CharSet.Auto)> _
	Public Shared Function GetWindowText(ByVal hWnd As Integer, _
	 ByVal lpString As StringBuilder, _
	 ByVal nMaxCount As Integer) As Integer
	End Function

	<DllImport("user32.dll", CharSet:=CharSet.Auto)> _
	Public Shared Function GetClassName(ByVal hwnd As Integer, _
	 ByVal lpClassName As StringBuilder, _
	 ByVal cch As Integer) As Integer
	End Function

	<DllImport("user32.dll", CharSet:=CharSet.Auto)> _
	Public Shared Function CloseWindow(ByVal hwnd As Integer) As Integer
	End Function

	<DllImport("user32.dll", CharSet:=CharSet.Auto)> _
	Public Shared Function SetForegroundWindow(ByVal hwnd As Integer) As Integer
	End Function

	<DllImport("user32.dll", CharSet:=CharSet.Auto)> _
	Public Shared Function GetForegroundWindow() As Integer
	End Function

	<DllImport("user32.dll", CharSet:=CharSet.Auto)> _
	Public Shared Function BringWindowToTop(ByVal hwnd As Integer) As Integer
	End Function

	<DllImport("user32.dll", CharSet:=CharSet.Auto)> _
	Public Shared Function AttachThreadInput(ByVal targte As Integer,ByVal current As Integer,ByVal flag As Integer) As Integer
	End Function

	<DllImport("user32.dll", CharSet:=CharSet.Auto)> _
	Public Shared Function GetWindowThreadProcessId(ByVal hwnd As Integer, proc As Integer) As Integer
	End Function

	Const STRING_BUFFER_LENGTH As Integer = 255

	Private Function CheckWindowAndActiveWindow1(ByVal hWnd As Integer, _
	  ByVal lParam As String) As Boolean

		Dim windowText As New StringBuilder(STRING_BUFFER_LENGTH)
		Dim className As New StringBuilder(STRING_BUFFER_LENGTH)
		Dim pt as integer
		Dim pc as integer
		Dim pn as integer

		GetWindowText(hWnd, windowText, STRING_BUFFER_LENGTH)
		GetClassName(hWnd, className, STRING_BUFFER_LENGTH)

		' for debug
'		System.Console.WriteLine(lParam)
'		System.Console.WriteLine(String.Format("{0}",windowText.ToString()))

		If windowText.ToString().contains("SCE-IT-SVR ") Then
			SetForegroundWindow(hWnd)
			pt=GetWindowThreadProcessId(hWnd,pn)
			pc=GetWindowThreadProcessId(GetForegroundWindow(),pn)
			If pt <> pc Then
				AttachThreadInput(pt,pc,1)
				BringWindowToTop(hWnd)
			End If
			count=1
		End If

		Return True

	End Function

	Private Function CheckWindowAndActiveWindow2(ByVal hWnd As Integer, _
	  ByVal lParam As String) As Boolean

		Dim windowText As New StringBuilder(STRING_BUFFER_LENGTH)
		Dim className As New StringBuilder(STRING_BUFFER_LENGTH)
		Dim pt as integer
		Dim pc as integer
		Dim pn as integer

		GetWindowText(hWnd, windowText, STRING_BUFFER_LENGTH)
		GetClassName(hWnd, className, STRING_BUFFER_LENGTH)

		' for debug
'		System.Console.WriteLine(lParam)
'		System.Console.WriteLine(String.Format("{0}",windowText.ToString()))

		If windowText.ToString().contains("TT10-P1100-10 ") Then
			SetForegroundWindow(hWnd)
			pt=GetWindowThreadProcessId(hWnd,pn)
			pc=GetWindowThreadProcessId(GetForegroundWindow(),pn)
			If pt <> pc Then
				AttachThreadInput(pt,pc,1)
				BringWindowToTop(hWnd)
			End If
			count=1
		End If

		Return True

	End Function

	Public Sub ActiveWindow1()

		count=0
		EnumWindowsString( New EnumWindowsCallbackString(AddressOf CheckWindowAndActiveWindow1), "")

	End Sub

	Public Sub ActiveWindow2()

		count=0
		EnumWindowsString( New EnumWindowsCallbackString(AddressOf CheckWindowAndActiveWindow2), "")

	End Sub

End Class
"@
Add-Type -TypeDefinition $source -Language VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$p = ps | ? {$_.Name -eq 'TT~P110010.V001.20170130.0000'}
if($p -eq $null){
	echo "no window!!"
	exit
}

echo "set focus to [退社]"

while(1){
	$t = Get-Date -UFormat "%H%M"
	if($t -eq "1705"){
		break
	}
	else{
		echo $t
	}
	sleep 10
}

$obj = New-Object API_EnumWindows

$ret = $obj.ActiveWindow1()
if($obj.count -eq 1){
	echo "[残業警告ウィンドウ]"
	sleep 1
	[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
}

$ret = $obj.ActiveWindow2()
if($obj.count -eq 1){
	sleep 1
	[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
	sleep(1)
	[System.Windows.Forms.SendKeys]::SendWait("{LEFT}{ENTER}")
	sleep(1)
	[System.Windows.Forms.SendKeys]::SendWait("{LEFT}{ENTER}")
	sleep(10)
	shutdown /s
}
