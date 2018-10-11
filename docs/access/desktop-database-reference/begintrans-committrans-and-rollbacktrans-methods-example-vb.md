---
title: BeginTrans, CommitTrans et RollbackTrans, méthodes-exemple (VB)
TOCTitle: BeginTrans, CommitTrans, and RollbackTrans Methods Example (VB)
ms:assetid: 12fce322-dba7-9159-8a09-7f6daf1a80ed
ms:mtpsurl: https://msdn.microsoft.com/library/JJ248904(v=office.15)
ms:contentKeyID: 48543357
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: b7a667130cbd3b756bfda9e73e54d83217600fcc
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471401"
---
# <a name="begintrans-committrans-and-rollbacktrans-methods-example-vb"></a><span data-ttu-id="5c4c9-102">BeginTrans, CommitTrans et RollbackTrans, méthodes - Exemple (VB)</span><span class="sxs-lookup"><span data-stu-id="5c4c9-102">BeginTrans, CommitTrans, and RollbackTrans Methods Example (VB)</span></span>


<span data-ttu-id="5c4c9-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="5c4c9-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="5c4c9-104">Cet exemple montre comment modifier le type de tous les livres de psychologie dans la table ***Titles*** de la base de données.</span><span class="sxs-lookup"><span data-stu-id="5c4c9-104">This example changes the book type of all psychology books in the ***Titles*** table of the database.</span></span> <span data-ttu-id="5c4c9-105">Une fois que la méthode [BeginTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) démarre une transaction qui isole toutes les modifications apportées à la table ***Titles*** , la méthode [CommitTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) enregistre les modifications.</span><span class="sxs-lookup"><span data-stu-id="5c4c9-105">After the [BeginTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) method starts a transaction that isolates all the changes made to the ***Titles*** table, the [CommitTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) method saves the changes.</span></span> <span data-ttu-id="5c4c9-106">Vous pouvez utiliser la méthode [RollbackTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) pour annuler les modifications que vous avez enregistré à l’aide de la méthode de [mise à jour](update-method-ado.md) .</span><span class="sxs-lookup"><span data-stu-id="5c4c9-106">You can use the [RollbackTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) method to undo changes that you saved using the [Update](update-method-ado.md) method.</span></span>

```vb 
 
'BeginBeginTransVB 
 
 'To integrate this code 
 'replace the data source and initial catalog values 
 'in the connection string 
 
Public Sub Main() 
 On Error GoTo ErrorHandler 
 
 'recordset and connection variables 
 Dim Cnxn As ADODB.Connection 
 Dim strCnxn As String 
 Dim rstTitles As ADODB.Recordset 
 Dim strSQLTitles As String 
 'record variables 
 Dim strTitle As String 
 Dim strMessage As String 
 
 ' Open connection 
 strCnxn = "Provider='sqloledb';Data Source='MySqlServer';" & _ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';" 
 Set Cnxn = New ADODB.Connection 
 Cnxn.Open strCnxn 
 
 ' Open recordset dynamic to allow for changes 
 Set rstTitles = New ADODB.Recordset 
 strSQLTitles = "Titles" 
 rstTitles.Open strSQLTitles, Cnxn, adOpenDynamic, adLockPessimistic, adCmdTable 
 
 Cnxn.BeginTrans 
 
 ' Loop through recordset and prompt user 
 ' to change the type for a specified title 
 
 rstTitles.MoveFirst 
 
 Do Until rstTitles.EOF 
 If Trim(rstTitles!Type) = "psychology" Then 
 strTitle = rstTitles!Title 
 strMessage = "Title: " & strTitle & vbCr & _ 
 "Change type to self help?" 
 
 ' If yes, change type for the specified title 
 If MsgBox(strMessage, vbYesNo) = vbYes Then 
 rstTitles!Type = "self_help" 
 rstTitles.Update 
 End If 
 End If 
 rstTitles.MoveNext 
 Loop 
 
 ' Prompt user to commit all changes made 
 If MsgBox("Save all changes?", vbYesNo) = vbYes Then 
 Cnxn.CommitTrans 
 Else 
 Cnxn.RollbackTrans 
 End If 
 
 ' Print recordset 
 rstTitles.Requery 
 rstTitles.MoveFirst 
 Do While Not rstTitles.EOF 
 Debug.Print rstTitles!Title & " - " & rstTitles!Type 
 rstTitles.MoveNext 
 Loop 
 
 ' Restore original data as this is a demo 
 rstTitles.MoveFirst 
 
 Do Until rstTitles.EOF 
 If Trim(rstTitles!Type) = "self_help" Then 
 rstTitles!Type = "psychology" 
 rstTitles.Update 
 End If 
 rstTitles.MoveNext 
 Loop 
 
 ' clean up 
 rstTitles.Close 
 Cnxn.Close 
 Set rstTitles = Nothing 
 Set Cnxn = Nothing 
 Exit Sub 
 
ErrorHandler: 
 ' clean up 
 If Not rstTitles Is Nothing Then 
 If rstTitles.State = adStateOpen Then rstTitles.Close 
 End If 
 Set rstTitles = Nothing 
 
 If Not Cnxn Is Nothing Then 
 If Cnxn.State = adStateOpen Then Cnxn.Close 
 End If 
 Set Cnxn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
End Sub 
 
 
'EndBeginTransVB 
```
