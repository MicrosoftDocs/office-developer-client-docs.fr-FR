---
title: Move, méthode - Exemple (VB)
TOCTitle: Move Method Example (VB)
ms:assetid: e3e837a4-f46d-f8ec-11ba-9b8d66105b13
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250161(v=office.15)
ms:contentKeyID: 48548328
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 120e21c670a3fa2b317d5276f32b8486b6a17a28
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470475"
---
# <a name="move-method-example-vb"></a><span data-ttu-id="860fe-102">Move, méthode - Exemple (VB)</span><span class="sxs-lookup"><span data-stu-id="860fe-102">Move Method Example (VB)</span></span>


<span data-ttu-id="860fe-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="860fe-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="860fe-104">Cet exemple de code montre comment utiliser la méthode [Move](move-method-ado.md) pour positionner le pointeur d'enregistrement en fonction de l'entrée de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="860fe-104">This example uses the [Move](move-method-ado.md) method to position the record pointer based on user input.</span></span>

```vb 
 
'BeginMoveVB 
 
 'To integrate this code 
 'replace the data source and initial catalog values 
 'in the connection string 
 
Public Sub Main() 
 On Error GoTo ErrorHandler 
 
 ' connection and recordset variables 
 Dim rstAuthors As ADODB.Recordset 
 Dim Cnxn As ADODB.Connection 
 Dim strCnxn As String 
 Dim strSQLAuthors As String 
 ' record variables 
 Dim varBookmark As Variant 
 Dim strCommand As String 
 Dim lngMove As Long 
 
 ' Open connection 
 Set Cnxn = New ADODB.Connection 
 strCnxn = "Provider='sqloledb';Data Source='MySqlServer';" & _ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';" 
 Cnxn.Open strCnxn 
 
 ' Open recordset from Authors table 
 Set rstAuthors = New ADODB.Recordset 
 rstAuthors.CursorLocation = adUseClient 
 ' Use client cursor to allow use of AbsolutePosition property 
 strSQLAuthors = "SELECT au_id, au_fname, au_lname, city, state FROM Authors ORDER BY au_lname" 
 rstAuthors.Open strSQLAuthors, strCnxn, adOpenStatic, adLockOptimistic, adCmdText 
 
 rstAuthors.MoveFirst 
 
 Do 
 ' Display information about current record and 
 ' ask how many records to move 
 
 strCommand = InputBox( _ 
 "Record " & rstAuthors.AbsolutePosition & _ 
 " of " & rstAuthors.RecordCount & vbCr & _ 
 "Author: " & rstAuthors!au_fname & _ 
 " " & rstAuthors!au_lname & vbCr & _ 
 "Location: " & rstAuthors!city & _ 
 ", " & rstAuthors!State & vbCr & vbCr & _ 
 "Enter number of records to Move " & _ 
 "(positive or negative).") 
 
 ' this is for exiting the loop 
 'lngMove = CLng(strCommand) 
 
 lngMove = CLng(Val(strCommand)) 
 If lngMove = 0 Then 
 MsgBox "You either entered a non-number or canceled the input box. Exit the application." 
 Exit Do 
 End If 
 
 ' Store bookmark in case the Move goes too far 
 ' forward or backward 
 varBookmark = rstAuthors.Bookmark 
 
 ' Move method requires parameter of data type Long 
 rstAuthors.Move lngMove 
 
 ' Trap for BOF or EOF 
 If rstAuthors.BOF Then 
 MsgBox "Too far backward! Returning to current record." 
 rstAuthors.Bookmark = varBookmark 
 End If 
 If rstAuthors.EOF Then 
 MsgBox "Too far forward! Returning to current record." 
 rstAuthors.Bookmark = varBookmark 
 End If 
 Loop 
 
 ' clean up 
 rstAuthors.Close 
 Cnxn.Close 
 Set rstAuthors = Nothing 
 Set Cnxn = Nothing 
 Exit Sub 
 
ErrorHandler: 
 ' clean up 
 If Not rstAuthors Is Nothing Then 
 If rstAuthors.State = adStateOpen Then rstAuthors.Close 
 End If 
 Set rstAuthors = Nothing 
 
 If Not Cnxn Is Nothing Then 
 If Cnxn.State = adStateOpen Then Cnxn.Close 
 End If 
 Set Cnxn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
End Sub 
'EndMoveVB 
```
