---
title: Database.RecordsAffected Property (DAO)
TOCTitle: RecordsAffected Property
ms:assetid: 1c591231-21dd-f0b1-4ba6-87784c5890d3
ms:mtpsurl: https://msdn.microsoft.com/library/Ff845732(v=office.15)
ms:contentKeyID: 48543567
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 8f0fb54773b6ab28a871b4a550e91dda5516c97c
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471979"
---
# <a name="databaserecordsaffected-property-dao"></a><span data-ttu-id="d9ce7-102">Database.RecordsAffected Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="d9ce7-102">Database.RecordsAffected Property (DAO)</span></span>


<span data-ttu-id="d9ce7-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="d9ce7-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="d9ce7-104">Renvoie le nombre d'enregistrements affectés par le dernier appel de la méthode **[Execute](connection-execute-method-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="d9ce7-104">Returns the number of records affected by the most recently invoked **[Execute](connection-execute-method-dao.md)** method.</span></span>

## <a name="syntax"></a><span data-ttu-id="d9ce7-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="d9ce7-105">Syntax</span></span>

<span data-ttu-id="d9ce7-106">*expression* . RecordsAffected</span><span class="sxs-lookup"><span data-stu-id="d9ce7-106">*expression* .RecordsAffected</span></span>

<span data-ttu-id="d9ce7-107">*expression* Variable qui représente un objet de **base de données** .</span><span class="sxs-lookup"><span data-stu-id="d9ce7-107">*expression* A variable that represents a **Database** object.</span></span>

## <a name="example"></a><span data-ttu-id="d9ce7-108">Exemple</span><span class="sxs-lookup"><span data-stu-id="d9ce7-108">Example</span></span>

<span data-ttu-id="d9ce7-p101">L'exemple ci-dessous utilise la propriété **RecordsAffected** avec des requêtes Action exécutées à partir d'objets **Database** et **QueryDef**. La fonction RecordsAffectedOutput est requise pour pouvoir exécuter cette procédure.</span><span class="sxs-lookup"><span data-stu-id="d9ce7-p101">This example uses the **RecordsAffected** property with action queries executed from a **Database** object and from a **QueryDef** object. The RecordsAffectedOutput function is required for this procedure to run.</span></span>

```vb
    Sub RecordsAffectedX() 
     
     Dim dbsNorthwind As Database 
     Dim qdfTemp As QueryDef 
     Dim strSQLChange As String 
     Dim strSQLRestore As String 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     
     With dbsNorthwind 
     ' Print report of contents of the Employees 
     ' table. 
     Debug.Print _ 
     "Number of records in Employees table: " & _ 
     .TableDefs!Employees.RecordCount 
     RecordsAffectedOutput dbsNorthwind 
     
     ' Define and execute an action query. 
     strSQLChange = "UPDATE Employees " & _ 
     "SET Country = 'United States' " & _ 
     "WHERE Country = 'USA'" 
     .Execute strSQLChange 
     
     ' Print report of contents of the Employees 
     ' table. 
     Debug.Print _ 
     "RecordsAffected after executing query " & _ 
     "from Database: " & .RecordsAffected 
     RecordsAffectedOutput dbsNorthwind 
     
     ' Define and run another action query. 
     strSQLRestore = "UPDATE Employees " & _ 
     "SET Country = 'USA' " & _ 
     "WHERE Country = 'United States'" 
     Set qdfTemp = .CreateQueryDef("", strSQLRestore) 
     qdfTemp.Execute 
     
     ' Print report of contents of the Employees 
     ' table. 
     Debug.Print _ 
     "RecordsAffected after executing query " & _ 
     "from QueryDef: " & qdfTemp.RecordsAffected 
     RecordsAffectedOutput dbsNorthwind 
     
     .Close 
     
     End With 
     
    End Sub 
     
    Function RecordsAffectedOutput(dbsNorthwind As Database) 
     
     Dim rstEmployees As Recordset 
     
     ' Open a Recordset object from the Employees table. 
     Set rstEmployees = _ 
     dbsNorthwind.OpenRecordset("Employees") 
     
     With rstEmployees 
     ' Enumerate Recordset. 
     .MoveFirst 
     Do While Not .EOF 
     Debug.Print " " & !LastName & ", " & !Country 
     .MoveNext 
     Loop 
     .Close 
     End With 
     
    End Function
```