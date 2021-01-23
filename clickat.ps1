# scheduledclick / clickat.ps1
# https://github.com/ttanimu/scheduledclick
# 
# Replace the following string before you run this.
#   <REPLACE TO WINDOW TITLE> : You can see this on the title bar of the window.
#   <REPLACE TO PROCESS NAME> : You can find this by running "ps" command on powershell.
#   <REPLACE TO TIME TO CLICK(HHMM)> : ex. 17:05 -> "1705"

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

	Private Function CheckWindowAndActiveCertainWindow(ByVal hWnd As Integer, _
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

		If windowText.ToString().contains("<REPLACE TO WINDOW TITLE>") Then
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

	Public Sub ActiveCertainWindow()

		count=0
		EnumWindowsString( New EnumWindowsCallbackString(AddressOf CheckWindowAndActiveCertainWindow), "")

	End Sub

End Class
"@
Add-Type -TypeDefinition $source -Language VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$p = ps | ? {$_.Name -eq '<REPLACE TO PROCESS NAME>'}
if($p -eq $null){
	echo "no window!!"
	exit
}

while(1){
	$t = Get-Date -UFormat "%H%M"
	if($t -eq "<REPLACE TO TIME TO CLICK(HHMM)>"){
		break
	}
	else{
		echo $t
	}
	sleep 10
}

$obj = New-Object API_EnumWindows

$ret = $obj.ActiveCertainWindow()
if($obj.count -eq 1){
	sleep 1
	[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
}
