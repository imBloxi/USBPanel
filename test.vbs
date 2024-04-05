Option Explicit

Dim objShell, objFSO, objNetwork, objComputer, strComputer, colComputers, strMessage
Dim objWMIService, colItems, objItem

' Create objects
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

' Get current computer name and IP address
strComputer = objNetwork.ComputerName
strComputer = Trim(strComputer)
strComputer = Replace(strComputer, " ", "")

Dim strIPAddress
strIPAddress = GetIPAddress()

' Construct IP range based on the current computer's IP
Dim arrIP
arrIP = Split(strIPAddress, ".")
Dim strIPRange
strIPRange = arrIP(0) & "." & arrIP(1) & "." & arrIP(2) & "."
WScript.Echo "Scanning network range: " & strIPRange & "*"

' Get list of computers on the network in the determined range
Set colComputers = objWMIService.ExecQuery _
    ("Select * From Win32_PingStatus Where Address Like '" & strIPRange & "%'")

' Loop through the list of computers
For Each objItem in colComputers
    ' Check if the computer is reachable
    If objItem.StatusCode = 0 Then
        strMessage = "Hello from " & strComputer & "!"
        
        ' Create a temporary VBS file to open notepad with the message
        Dim tempVBSFile
        Set tempVBSFile = objFSO.CreateTextFile("C:\Temp\OpenNotepad.vbs")
        tempVBSFile.WriteLine "Set objShell = CreateObject(""WScript.Shell"")"
        tempVBSFile.WriteLine "objShell.Run ""notepad.exe"""
        tempVBSFile.WriteLine "WScript.Sleep 1000"
        tempVBSFile.WriteLine "objShell.SendKeys """ & strMessage & """"
        tempVBSFile.Close
        
        ' Execute the temporary VBS file
        objShell.Run "C:\Temp\OpenNotepad.vbs", 1, False
        
        ' Wait for notepad to open and display the message
        WScript.Sleep 3000
        
        ' Delete the temporary VBS file
        objFSO.DeleteFile "C:\Temp\OpenNotepad.vbs"
    End If
Next

' Clean up objects
Set objShell = Nothing
Set objFSO = Nothing
Set objNetwork = Nothing
Set objWMIService = Nothing

' Function to get the IP Address of the local computer
Function GetIPAddress()
    Dim objWMIService, colItems, objItem

    ' Get WMI object
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

    ' Query for IP Address
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

    ' Loop through the items (usually just one)
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then
            GetIPAddress = objItem.IPAddress(0)
            Exit Function
        End If
    Next
End Function