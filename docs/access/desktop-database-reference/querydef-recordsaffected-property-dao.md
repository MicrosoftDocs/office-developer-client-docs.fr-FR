---
title: QueryDef.RecordsAffected Property (DAO)
TOCTitle: RecordsAffected Property
ms:assetid: 29a864b5-305c-d33f-b2ca-fc9a08baaa5c
ms:mtpsurl: https://msdn.microsoft.com/library/Ff192040(v=office.15)
ms:contentKeyID: 48543886
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1053082
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 9377f61ec334f41b7f8ccba61b87a9d5953b4429
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471397"
---
# <a name="querydefrecordsaffected-property-dao"></a><span data-ttu-id="5e158-102">QueryDef.RecordsAffected Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="5e158-102">QueryDef.RecordsAffected Property (DAO)</span></span>


<span data-ttu-id="5e158-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="5e158-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="5e158-104">Renvoie le nombre d'enregistrements affectés par le dernier appel de la méthode **[Execute](querydef-execute-method-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="5e158-104">Returns the number of records affected by the most recently invoked **[Execute](querydef-execute-method-dao.md)** method.</span></span>

## <a name="syntax"></a><span data-ttu-id="5e158-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="5e158-105">Syntax</span></span>

<span data-ttu-id="5e158-106">*expression* . RecordsAffected</span><span class="sxs-lookup"><span data-stu-id="5e158-106">*expression* .RecordsAffected</span></span>

<span data-ttu-id="5e158-107">*expression* Variable qui représente un objet **QueryDef** .</span><span class="sxs-lookup"><span data-stu-id="5e158-107">*expression* A variable that represents a **QueryDef** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="5e158-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="5e158-108">Remarks</span></span>

<span data-ttu-id="5e158-109">Lorsque vous utilisez la méthode **Execute** pour exécuter une requête Action à partir d'un objet **QueryDef**, la propriété **RecordsAffected** contiendra le nombre d'enregistrements supprimés, mis à jour ou insérés.</span><span class="sxs-lookup"><span data-stu-id="5e158-109">When you use the **Execute** method to run an action query from a **QueryDef** object, the **RecordsAffected** property will contain the number of records deleted, updated, or inserted.</span></span>

## <a name="example"></a><span data-ttu-id="5e158-110">Exemple</span><span class="sxs-lookup"><span data-stu-id="5e158-110">Example</span></span>

<span data-ttu-id="5e158-p101">Cet exemple utilise la propriété **RecordsAffected** avec des requêtes action exécutées à partir d'un objet **[Database](database-object-dao.md)** et d'un objet **QueryDef**. La fonction RecordsAffectedOutput est nécessaire à l'exécution de cette procédure.</span><span class="sxs-lookup"><span data-stu-id="5e158-p101">This example uses the **RecordsAffected** property with action queries executed from a **[Database](database-object-dao.md)** object and from a **QueryDef** object. The RecordsAffectedOutput function is required for this procedure to run.</span></span>

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