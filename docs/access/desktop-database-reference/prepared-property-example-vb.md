---
title: Prepared, propriété – Exemple (VB)
TOCTitle: Prepared property example (VB)
ms:assetid: d7332052-bf2e-f7d4-eb06-59ff8d68f812
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250081(v=office.15)
ms:contentKeyID: 48548000
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: a4d66edef0ca8a789c912d3c4d00c57f4a17ef36
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32301448"
---
# <a name="prepared-property-example-vb"></a>Prepared, propriété – Exemple (VB)


**S’applique à** : Access 2013, Office 2013

Cet exemple illustre la propriété [Prepared](prepared-property-ado.md) par l’ouverture de deux objets [Command](command-object-ado.md) : l’un préparé et l’autre non préparé.

```vb 
 
'BeginPreparedVB 
 
 'To integrate this code 
 'replace the data source and initial catalog values 
 'in the connection string 
 
Public Sub Main() 
 On Error GoTo ErrorHandler 
 
 Dim Cnxn As ADODB.Connection 
 Dim cmd1 As ADODB.Command 
 Dim cmd2 As ADODB.Command 
 
 Dim strCnxn As String 
 Dim strCmd As String 
 Dim sngStart As Single 
 Dim sngEnd As Single 
 Dim sngNotPrepared As Single 
 Dim sngPrepared As Single 
 Dim intLoop As Integer 
 
 ' Open a connection 
 strCnxn = "Provider='sqloledb';Data Source='MySqlServer';" & _ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';" 
 Set Cnxn = New ADODB.Connection 
 Cnxn.Open strCnxn 
 
 ' Create two command objects for the same 
 ' command - one prepared and one not prepared 
 strCmd = "SELECT title, type FROM Titles ORDER BY type" 
 
 Set cmd1 = New ADODB.Command 
 Set cmd1.ActiveConnection = Cnxn 
 cmd1.CommandText = strCmd 
 
 Set cmd2 = New ADODB.Command 
 Set cmd2.ActiveConnection = Cnxn 
 cmd2.CommandText = strCmd 
 cmd2.Prepared = True 
 
 ' Set a timer, then execute the unprepared 
 ' command 20 times 
 sngStart = Timer 
 For intLoop = 1 To 20 
 cmd1.Execute 
 Next intLoop 
 sngEnd = Timer 
 sngNotPrepared = sngEnd - sngStart 
 
 ' Reset the timer, then execute the prepared 
 ' command 20 times 
 sngStart = Timer 
 For intLoop = 1 To 20 
 cmd2.Execute 
 Next intLoop 
 sngEnd = Timer 
 sngPrepared = sngEnd - sngStart 
 
 ' Display performance results 
 MsgBox "Performance Results:" & vbCr & _ 
 " Not Prepared: " & Format(sngNotPrepared, _ 
 "##0.000") & " seconds" & vbCr & _ 
 " Prepared: " & Format(sngPrepared, _ 
 "##0.000") & " seconds" 
 
 ' clean up 
 Cnxn.Close 
 Set Cnxn = Nothing 
 Set cmd1 = Nothing 
 Set cmd2 = Nothing 
 Exit Sub 
 
ErrorHandler: 
 ' clean up 
 Set cmd1 = Nothing 
 Set cmd2 = Nothing 
 
 If Not Cnxn Is Nothing Then 
 If Cnxn.State = adStateOpen Then Cnxn.Close 
 End If 
 Set Cnxn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
End Sub 
'EndPreparedVB 
```

