---
title: CursorType, LockType et EditMode, propriétés - Exemple (VB)
TOCTitle: CursorType, LockType, and EditMode Properties Example (VB)
ms:assetid: efe3f976-b095-c0ce-376a-693b07ec8e9d
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250221(v=office.15)
ms:contentKeyID: 48548595
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: aba88fbbf73801b99dcb7adb20639d8c1ae06641
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472215"
---
# <a name="cursortype-locktype-and-editmode-properties-example-vb"></a><span data-ttu-id="a2062-102">CursorType, LockType et EditMode, propriétés - Exemple (VB)</span><span class="sxs-lookup"><span data-stu-id="a2062-102">CursorType, LockType, and EditMode Properties Example (VB)</span></span>


<span data-ttu-id="a2062-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="a2062-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="a2062-p101">Cet exemple illustre la définition des propriétés [CursorType](cursortype-property-ado.md) et [LockType](locktype-property-ado.md) avant l'ouverture d'un objet [Recordset](recordset-object-ado.md). Il indique également la valeur de la propriété [EditMode](editmode-property-ado.md) dans différentes conditions. La fonction **EditModeOutput** est nécessaire à l'exécution de cette procédure.</span><span class="sxs-lookup"><span data-stu-id="a2062-p101">This example demonstrates setting the [CursorType](cursortype-property-ado.md) and [LockType](locktype-property-ado.md) properties before opening a [Recordset](recordset-object-ado.md). It also shows the value of the [EditMode](editmode-property-ado.md) property under various conditions. The **EditModeOutput** function is required for this procedure to run.</span></span>

```vb 
 
'BeginEditModeVB 
 
 'To integrate this code 
 'replace the data source and initial catalog values 
 'in the connection string 
 
Public Sub Main() 
 On Error GoTo ErrorHandler 
 
 ' recordset variables 
 Dim rstEmployees As ADODB.Recordset 
 Dim Cnxn As ADODB.Connection 
 Dim strCnxn As String 
 Dim SQLEmployees As String 
 
 ' open connection 
 Set Cnxn = New ADODB.Connection 
 strCnxn = "Provider='sqloledb';Data Source='MySqlServer';" & _ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';" 
 Cnxn.Open strCnxn 
 
 ' set recordset properties through object refs 
 ' instead of through arguments to Open method 
 Set rstEmployees = New ADODB.Recordset 
 Set rstEmployees.ActiveConnection = Cnxn 
 rstEmployees.CursorLocation = adUseClient 
 rstEmployees.CursorType = adOpenStatic 
 rstEmployees.LockType = adLockBatchOptimistic 
 
 ' open recordset with data from Employee table 
 SQLEmployees = "employee" 
 rstEmployees.Open SQLEmployees, , , , adCmdTable 
 
 ' Show the EditMode property under different editing states 
 rstEmployees.AddNew 
 rstEmployees!emp_id = "T-T55555M" 
 rstEmployees!fname = "temp_fname" 
 rstEmployees!lname = "temp_lname" 
 'call function below 
 'to output results to debug window 
 EditModeOutput "After AddNew:", rstEmployees.EditMode 
 rstEmployees.UpdateBatch 
 EditModeOutput "After UpdateBatch:", rstEmployees.EditMode 
 rstEmployees!fname = "test" 
 EditModeOutput "After Edit:", rstEmployees.EditMode 
 rstEmployees.Close 
 
 ' Delete new record because this is a demonstration 
 Cnxn.Execute "DELETE FROM employee WHERE emp_id = 'T-T55555M'" 
 
 ' clean up 
 Cnxn.Close 
 Set rstEmployees = Nothing 
 Set Cnxn = Nothing 
 Exit Sub 
 
ErrorHandler: 
 ' clean up 
 If Not rstEmployees Is Nothing Then 
 If rstEmployees.State = adStateOpen Then rstEmployees.Close 
 End If 
 Set rstEmployees = Nothing 
 
 If Not Cnxn Is Nothing Then 
 If Cnxn.State = adStateOpen Then Cnxn.Close 
 End If 
 Set Cnxn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
 
End Sub 
 
Public Function EditModeOutput(strTemp As String, _ 
 intEditMode As Integer) 
 
 ' Print report based on the value of the EditMode 
 ' property 
 Debug.Print strTemp 
 Debug.Print " EditMode = "; 
 
 Select Case intEditMode 
 Case adEditNone 
 Debug.Print "adEditNone" 
 Case adEditInProgress 
 Debug.Print "adEditInProgress" 
 Case adEditAdd 
 Debug.Print "adEditAdd" 
 End Select 
 
End Function 
'EndEditModeVB 
```
