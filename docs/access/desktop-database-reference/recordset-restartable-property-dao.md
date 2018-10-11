---
title: Recordset.Restartable Property (DAO)
TOCTitle: Restartable Property
ms:assetid: 00def49d-ea7e-6cd5-2f4a-914a1ddcdd51
ms:mtpsurl: https://msdn.microsoft.com/library/Ff844737(v=office.15)
ms:contentKeyID: 48542919
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052926
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 2b36a2dd7400b2d3376149a4d012c8fa79242021
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470679"
---
# <a name="recordsetrestartable-property-dao"></a><span data-ttu-id="79fbd-102">Recordset.Restartable Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="79fbd-102">Recordset.Restartable Property (DAO)</span></span>


<span data-ttu-id="79fbd-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="79fbd-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="79fbd-104">Renvoie une valeur indiquant si un objet **[Recordset](recordset-object-dao.md)** prend en charge la méthode **[Requery](recordset-requery-method-dao.md)**, laquelle réexécute la requête sur laquelle l'objet **Recordset** est basé.</span><span class="sxs-lookup"><span data-stu-id="79fbd-104">Returns a value that indicates whether a **[Recordset](recordset-object-dao.md)** object supports the **[Requery](recordset-requery-method-dao.md)** method, which re-executes the query on which the **Recordset** object is based.</span></span>

## <a name="syntax"></a><span data-ttu-id="79fbd-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="79fbd-105">Syntax</span></span>

<span data-ttu-id="79fbd-106">*expression* . Restartable</span><span class="sxs-lookup"><span data-stu-id="79fbd-106">*expression* .Restartable</span></span>

<span data-ttu-id="79fbd-107">*expression* Variable qui représente un objet **Recordset** .</span><span class="sxs-lookup"><span data-stu-id="79fbd-107">*expression* A variable that represents a **Recordset** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="79fbd-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="79fbd-108">Remarks</span></span>

<span data-ttu-id="79fbd-109">Les objets **Recordset** de type table renvoient toujours la valeur **False**.</span><span class="sxs-lookup"><span data-stu-id="79fbd-109">Table-type **Recordset** objects always return **False**.</span></span>

<span data-ttu-id="79fbd-p101">Vérifiez la propriété **Restartable** avant d'utiliser la méthode **Requery** sur un objet **Recordset**. Si la propriété **Restartable** a la valeur **False**, appelez la méthode **[OpenRecordset](connection-openrecordset-method-dao.md)** sur l'objet **[QueryDef](querydef-object-dao.md)** sous-jacent pour réexécuter la requête.</span><span class="sxs-lookup"><span data-stu-id="79fbd-p101">Check the **Restartable** property before using the **Requery** method on a **Recordset** object. If the object's **Restartable** property is set to **False**, use the **[OpenRecordset](connection-openrecordset-method-dao.md)** method on the underlying **[QueryDef](querydef-object-dao.md)** object to re-execute the query.</span></span>

## <a name="example"></a><span data-ttu-id="79fbd-112">Exemple</span><span class="sxs-lookup"><span data-stu-id="79fbd-112">Example</span></span>

<span data-ttu-id="79fbd-113">Cet exemple illustre l'utilisation de la propriété **Restartable** avec différents objets **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="79fbd-113">This example demonstrates the **Restartable** property with different **Recordset** objects.</span></span>

```vb
    Sub RestartableX() 
     
     Dim dbsNorthwind As Database 
     Dim rstTemp As Recordset 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     
     With dbsNorthwind 
     ' Open a table-type Recordset and print its 
     ' Restartable property. 
     Set rstTemp = .OpenRecordset("Employees", dbOpenTable) 
     Debug.Print _ 
     "Table-type recordset from Employees table" 
     Debug.Print " Restartable = " & rstTemp.Restartable 
     rstTemp.Close 
     
     ' Open a Recordset from an SQL statement and print its 
     ' Restartable property. 
     Set rstTemp = _ 
     .OpenRecordset("SELECT * FROM Employees") 
     Debug.Print "Recordset based on SQL statement" 
     Debug.Print " Restartable = " & rstTemp.Restartable 
     rstTemp.Close 
     
     ' Open a Recordset from a saved QueryDef object and 
     ' print its Restartable property. 
     Set rstTemp = .OpenRecordset("Current Product List") 
     Debug.Print _ 
     "Recordset based on permanent QueryDef (" & _ 
     rstTemp.Name & ")" 
     Debug.Print " Restartable = " & rstTemp.Restartable 
     rstTemp.Close 
     
     .Close 
     End With 
     
    End Sub
```