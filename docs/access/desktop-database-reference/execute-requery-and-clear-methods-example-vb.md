---
title: Execute, Requery et Clear, méthodes - Exemple (VB)
TOCTitle: Execute, Requery, and Clear Methods Example (VB)
ms:assetid: 6d700971-6b77-bd41-dd22-df53f902c0f2
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249432(v=office.15)
ms:contentKeyID: 48545491
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 6d32d4e0c56ea4b6a03474562410255bbc5952f4
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471456"
---
# <a name="execute-requery-and-clear-methods-example-vb"></a><span data-ttu-id="9160a-102">Execute, Requery et Clear, méthodes - Exemple (VB)</span><span class="sxs-lookup"><span data-stu-id="9160a-102">Execute, Requery, and Clear Methods Example (VB)</span></span>


<span data-ttu-id="9160a-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="9160a-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="9160a-p101">Cet exemple illustre la méthode **Execute** lorsqu'elle est exécutée à la fois à partir d'un objet [Command](command-object-ado.md) et d'un objet [Connection](connection-object-ado.md). Elle utilise également la méthode [Requery](requery-method-ado.md) pour récupérer les données actuelles d'un objet [Recordset](recordset-object-ado.md) et la méthode [Clear](clear-method-ado.md) pour effacer le contenu de la collection [Errors](errors-collection-ado.md). (L'accès à la collection **Errors** est réalisé via l'objet **Connection** de la propriété [ActiveConnection](activeconnection-property-ado.md) de l'objet [Recordset](recordset-object-ado.md).) Les procédures ExecuteCommand et PrintOutput sont nécessaires à l'exécution de cette procédure.</span><span class="sxs-lookup"><span data-stu-id="9160a-p101">This example demonstrates the **Execute** method when run from both a [Command](command-object-ado.md) object and a [Connection](connection-object-ado.md) object. It also uses the [Requery](requery-method-ado.md) method to retrieve current data in a [Recordset](recordset-object-ado.md), and the [Clear](clear-method-ado.md) method to clear the contents of the [Errors](errors-collection-ado.md) collection. (The **Errors** collection is accessed via the **Connection** object of the [ActiveConnection](activeconnection-property-ado.md) property of the [Recordset](recordset-object-ado.md).) The ExecuteCommand and PrintOutput procedures are required for this procedure to run.</span></span>

```vb 
 
'BeginExecuteVB 
 
 'To integrate this code 
 'replace the data source and initial catalog values 
 'in the connection string 
 
Public Sub Main() 
 On Error GoTo Err_Execute 
 
 ' connection, command, and recordset variables 
 Dim Cnxn As ADODB.Connection 
 Dim cmdChange As ADODB.Command 
 Dim rstTitles As ADODB.Recordset 
 Dim Err As ADODB.Error 
 Dim strSQLChange As String 
 Dim strSQLRestore As String 
 Dim strSQLTitles 
 Dim strCnxn As String 
 
 ' Define two SQL statements to execute as command text 
 strSQLChange = "UPDATE Titles SET Type = 'self_help' WHERE Type = 'psychology'" 
 strSQLRestore = "UPDATE Titles SET Type = 'psychology' WHERE Type = 'self_help'" 
 
 ' Open connection 
 strCnxn = "Provider='sqloledb';Data Source='MySqlServer';" & _ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';" 
 Set Cnxn = New ADODB.Connection 
 Cnxn.Open strCnxn 
 
 ' Create command object 
 Set cmdChange = New ADODB.Command 
 Set cmdChange.ActiveConnection = Cnxn 
 cmdChange.CommandText = strSQLChange 
 
 ' Open titles table 
 Set rstTitles = New ADODB.Recordset 
 strSQLTitles = "titles" 
 rstTitles.Open strSQLTitles, Cnxn, , , adCmdTable 
 
 ' Print report of original data 
 Debug.Print _ 
 "Data in Titles table before executing the query" 
 PrintOutput rstTitles 
 
 ' Call the ExecuteCommand subroutine below to execute cmdChange command 
 ExecuteCommand cmdChange, rstTitles 
 
 ' Print report of new data 
 Debug.Print _ 
 "Data in Titles table after executing the query" 
 PrintOutput rstTitles 
 
 ' Use the Connection object's execute method to 
 ' execute SQL statement to restore data and trap for 
 ' errors, checking the Errors collection if necessary 
 Cnxn.Execute strSQLRestore, , adExecuteNoRecords 
 
 ' Retrieve the current data by requerying the recordset 
 rstTitles.Requery 
 
 ' Print report of restored data using sub from below 
 Debug.Print "Data after executing the query to restore the original information " 
 PrintOutput rstTitles 
 
 ' clean up 
 rstTitles.Close 
 Cnxn.Close 
 Set rstTitles = Nothing 
 Set Cnxn = Nothing 
 Exit Sub 
 
Err_Execute: 
 ' Notify user of any errors that result from 
 ' executing the query 
 If rstTitles.ActiveConnection.Errors.Count >= 0 Then 
 For Each Err In rstTitles.ActiveConnection.Errors 
 MsgBox "Error number: " & Err.Number & vbCr & _ 
 Err.Description 
 Next Err 
 End If 
 
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
 
Public Sub ExecuteCommand(cmdTemp As ADODB.Command, rstTemp As ADODB.Recordset) 
 
 Dim Err As Error 
 
 ' Run the specified Command object and trap for 
 ' errors, checking the Errors collection 
 On Error GoTo Err_Execute 
 cmdTemp.Execute 
 On Error GoTo 0 
 
 ' Retrieve the current data by requerying the recordset 
 rstTemp.Requery 
 
 Exit Sub 
 
Err_Execute: 
 
 ' Notify user of any errors that result from 
 ' executing the query 
 If rstTemp.ActiveConnection.Errors.Count > 0 Then 
 For Each Err In rstTemp.ActiveConnection.Errors 
 MsgBox "Error number: " & Err.Number & vbCr & _ 
 Err.Description 
 Next Err 
 End If 
 
 Resume Next 
 
End Sub 
 
Public Sub PrintOutput(rstTemp As ADODB.Recordset) 
 
 ' Enumerate Recordset 
 Do While Not rstTemp.EOF 
 Debug.Print " " & rstTemp!Title & _ 
 ", " & rstTemp!Type 
 rstTemp.MoveNext 
 Loop 
 
End Sub 
'EndExecuteVB 
```
