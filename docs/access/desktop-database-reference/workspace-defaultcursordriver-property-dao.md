---
title: Workspace.DefaultCursorDriver Property (DAO)
TOCTitle: DefaultCursorDriver Property
ms:assetid: 15a8356d-7ae0-3c8e-fbb7-2d8ad6d9a582
ms:mtpsurl: https://msdn.microsoft.com/library/Ff845499(v=office.15)
ms:contentKeyID: 48543413
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1053582
f1_categories:
- Office.Version=v15
ms.openlocfilehash: fa2e25228d5536e13a4fcff8af329708fd9a90e7
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471056"
---
# <a name="workspacedefaultcursordriver-property-dao"></a><span data-ttu-id="0e0ed-102">Workspace.DefaultCursorDriver Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="0e0ed-102">Workspace.DefaultCursorDriver Property (DAO)</span></span>


<span data-ttu-id="0e0ed-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="0e0ed-103">**Applies to**: Access 2013 | Office 2013</span></span>


## <a name="syntax"></a><span data-ttu-id="0e0ed-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="0e0ed-104">Syntax</span></span>

<span data-ttu-id="0e0ed-105">*expression* . DefaultCursorDriver</span><span class="sxs-lookup"><span data-stu-id="0e0ed-105">*expression* .DefaultCursorDriver</span></span>

<span data-ttu-id="0e0ed-106">*expression* Variable qui représente un objet **Workspace** .</span><span class="sxs-lookup"><span data-stu-id="0e0ed-106">*expression* A variable that represents a **Workspace** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="0e0ed-107">Remarques</span><span class="sxs-lookup"><span data-stu-id="0e0ed-107">Remarks</span></span>

<span data-ttu-id="0e0ed-108">Le paramètre ou la valeur renvoyée peut correspondre à l'une des constantes **[CursorDriverEnum](cursordriverenum-enumeration-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="0e0ed-108">The setting or return value can be set to one of the **[CursorDriverEnum](cursordriverenum-enumeration-dao.md)** constants.</span></span>

<span data-ttu-id="0e0ed-p101">Le paramètre de cette propriété ne concerne que les connexions établies après la définition de la propriété. En effet, la modification de la propriété **DefaultCursorDriver** n'a aucune conséquence sur les connexions existantes.</span><span class="sxs-lookup"><span data-stu-id="0e0ed-p101">This property setting only affects connections established after the property has been set. Changing the **DefaultCursorDriver** property has no effect on existing connections.</span></span>

## <a name="example"></a><span data-ttu-id="0e0ed-111">Exemple</span><span class="sxs-lookup"><span data-stu-id="0e0ed-111">Example</span></span>

<span data-ttu-id="0e0ed-p102">L'exemple ci-dessous utilise la méthode **NextRecordset** pour afficher les données à partir d'une requête SELECT composée. La propriété **DefaultCursorDriver** doit être définie sur la valeur **dbUseODBCCursor** en cas d'exécution de ce type de requête. La méthode **NextRecordset** renvoie la valeur **True** même si une partie ou l'ensemble des instructions SELECT renvoient des enregistrements nuls. Elle ne renvoie la valeur **False** qu'une fois toutes les clauses SQL vérifiées.</span><span class="sxs-lookup"><span data-stu-id="0e0ed-p102">This example uses the **NextRecordset** method to view the data from a compound SELECT query. The **DefaultCursorDriver** property must be set to **dbUseODBCCursor** when executing such queries. The **NextRecordset** method will return **True** even if some or all of the SELECT statements return zero records; it will return **False** only after all the individual SQL clauses have been checked.</span></span>

```vb
    Sub NextRecordsetX() 
     
     Dim wrkODBC As Workspace 
     Dim conPubs As Connection 
     Dim rstTemp As Recordset 
     Dim intCount As Integer 
     Dim booNext As Boolean 
     
     ' Create ODBCDirect Workspace object and open Connection 
     ' object. The DefaultCursorDriver setting is required 
     ' when using compound SQL statements. 
     Set wrkODBC = CreateWorkspace("", _ 
     "admin", "", dbUseODBC) 
     wrkODBC.DefaultCursorDriver = dbUseODBCCursor 
     
     ' Note: The DSN referenced below must be set to 
     ' use Microsoft Windows NT Authentication Mode to 
     ' authorize user access to the Microsoft SQL Server. 
     Set conPubs = wrkODBC.OpenConnection("Publishers", , , _ 
     "ODBC;DATABASE=pubs;DSN=Publishers") 
     
     ' Construct compound SELECT statement. 
     Set rstTemp = conPubs.OpenRecordset("SELECT * " & _ 
     "FROM authors; " & _ 
     "SELECT * FROM stores; " & _ 
     "SELECT * FROM jobs") 
     
     ' Try printing results from each of the three SELECT 
     ' statements. 
     booNext = True 
     intCount = 1 
     With rstTemp 
     Do While booNext 
     Debug.Print "Contents of recordset #" & intCount 
     Do While Not .EOF 
     Debug.Print , .Fields(0), .Fields(1) 
     .MoveNext 
     Loop 
     booNext = .NextRecordset 
     Debug.Print " rstTemp.NextRecordset = " & _ 
     booNext 
     intCount = intCount + 1 
     Loop 
     End With 
     
     rstTemp.Close 
     conPubs.Close 
     wrkODBC.Close 
     
    End Sub 
```

<br/>

<span data-ttu-id="0e0ed-p103">Pour effectuer cette tâche, il est également possible de créer une instruction préparée qui contient l'instruction SQL composée. La propriété **CacheSize** de l'objet **QueryDef** doit être définie sur la valeur 1 et l'objet **Recordset** doit être de type avant uniquement et en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="0e0ed-p103">Another way to accomplish the same task would be to create a prepared statement containing the compound SQL statement. The **CacheSize** property of the **QueryDef** object must be set to 1, and the **Recordset** object must be forward-only and read-only.</span></span>

```vb 
Sub NextRecordsetX2() 
 
 Dim wrkODBC As Workspace 
 Dim conPubs As Connection 
 Dim qdfTemp As QueryDef 
 Dim rstTemp As Recordset 
 Dim intCount As Integer 
 Dim booNext As Boolean 
 
 ' Create ODBCDirect Workspace object and open Connection 
 ' object. The DefaultCursorDriver setting is required 
 ' when using compound SQL statements. 
 Set wrkODBC = CreateWorkspace("", _ 
 "admin", "", dbUseODBC) 
 wrkODBC.DefaultCursorDriver = dbUseODBCCursor 
 
 ' Note: The DSN referenced below must be set to 
 ' use Microsoft Windows NT Authentication Mode to 
 ' authorize user access to the Microsoft SQL Server. 
 Set conPubs = wrkODBC.OpenConnection("Publishers", , , _ 
 "ODBC;DATABASE=pubs;DSN=Publishers") 
 
 ' Create a temporary stored procedure with a compound 
 ' SELECT statement. 
 Set qdfTemp = conPubs.CreateQueryDef("", _ 
 "SELECT * FROM authors; " & _ 
 "SELECT * FROM stores; " & _ 
 "SELECT * FROM jobs") 
 ' Set CacheSize and open Recordset object with arguments 
 ' that will allow access to multiple recordsets. 
 qdfTemp.CacheSize = 1 
 Set rstTemp = qdfTemp.OpenRecordset(dbOpenForwardOnly, _ 
 dbReadOnly) 
 
 ' Try printing results from each of the three SELECT 
 ' statements. 
 booNext = True 
 intCount = 1 
 With rstTemp 
 Do While booNext 
 Debug.Print "Contents of recordset #" & intCount 
 Do While Not .EOF 
 Debug.Print , .Fields(0), .Fields(1) 
 .MoveNext 
 Loop 
 booNext = .NextRecordset 
 Debug.Print " rstTemp.NextRecordset = " & _ 
 booNext 
 intCount = intCount + 1 
 Loop 
 End With 
 
 rstTemp.Close 
 qdfTemp.Close 
 conPubs.Close 
 wrkODBC.Close 
 
End Sub 
 
```

<br/>

<span data-ttu-id="0e0ed-p104">L'exemple ci-dessous fait appel aux propriétés **RecordStatus** et **DefaultCursorDriver** pour indiquer comment le suivi des modifications apportées à un objet **Recordset** local est effectué lors d'une mise à jour par lot. La fonction RecordStatusOutput est requise pour pouvoir exécuter cette procédure.</span><span class="sxs-lookup"><span data-stu-id="0e0ed-p104">This example uses the **RecordStatus** and **DefaultCursorDriver** properties to show how changes to a local **Recordset** are tracked during batch updating. The RecordStatusOutput function is required for this procedure to run.</span></span>

```vb 
Sub RecordStatusX() 
 
 Dim wrkMain As Workspace 
 Dim conMain As Connection 
 Dim rstTemp As Recordset 
 
 Set wrkMain = CreateWorkspace("ODBCWorkspace", _ 
 "admin", "", dbUseODBC) 
 ' This DefaultCursorDriver setting is required for 
 ' batch updating. 
 wrkMain.DefaultCursorDriver = dbUseClientBatchCursor 
 
 ' Note: The DSN referenced below must be configured to 
 ' use Microsoft Windows NT Authentication Mode to 
 ' authorize user access to the Microsoft SQL Server. 
 Set conMain = wrkMain.OpenConnection("Publishers", _ 
 dbDriverNoPrompt, False, _ 
 "ODBC;DATABASE=pubs;DSN=Publishers") 
 
 ' The following locking argument is required for 
 ' batch updating. 
 Set rstTemp = conMain.OpenRecordset( _ 
 "SELECT * FROM authors", dbOpenDynaset, 0, _ 
 dbOptimisticBatch) 
 
 With rstTemp 
 .MoveFirst 
 Debug.Print "Original record: " & !au_lname 
 Debug.Print , RecordStatusOutput2(.RecordStatus) 
 
 .Edit 
 !au_lname = "Bowen" 
 .Update 
 Debug.Print "Edited record: " & !au_lname 
 Debug.Print , RecordStatusOutput2(.RecordStatus) 
 
 .AddNew 
 !au_lname = "NewName" 
 .Update 
 Debug.Print "New record: " & !au_lname 
 Debug.Print , RecordStatusOutput2(.RecordStatus) 
 
 .Delete 
 Debug.Print "Deleted record: " & !au_lname 
 Debug.Print , RecordStatusOutput2(.RecordStatus) 
 
 ' Close the local recordset without updating the 
 ' data on the server. 
 .Close 
 End With 
 
 conMain.Close 
 wrkMain.Close 
 
End Sub 
 
Function RecordStatusOutput(lngTemp As Long) As String 
 
 Dim strTemp As String 
 
 strTemp = "" 
 
 ' Construct an output string based on the RecordStatus 
 ' value. 
If lngTemp = dbRecordUnmodified Then _ 
 strTemp = "[dbRecordUnmodified]" 
 If lngTemp = dbRecordModified Then _ 
 strTemp = "[dbRecordModified]" 
 If lngTemp = dbRecordNew Then _ 
 strTemp = "[dbRecordNew]" 
 If lngTemp = dbRecordDeleted Then _ 
 strTemp = "[dbRecordDeleted]" 
 If lngTemp = dbRecordDBDeleted Then _ 
 strTemp = "[dbRecordDBDeleted]" 
 
 RecordStatusOutput = strTemp 
 
End Function 
 
```
