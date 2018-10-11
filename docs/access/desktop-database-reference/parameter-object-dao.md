---
title: Parameter Object (DAO)
TOCTitle: Parameter Object
ms:assetid: 194efd23-6086-13ac-beb9-c2aec101d6fe
ms:mtpsurl: https://msdn.microsoft.com/library/Ff845640(v=office.15)
ms:contentKeyID: 48543495
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 69388834d5469c9fa66d70d397d63be4376db218
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472275"
---
# <a name="parameter-object-dao"></a><span data-ttu-id="addfc-102">Parameter Object (DAO)</span><span class="sxs-lookup"><span data-stu-id="addfc-102">Parameter Object (DAO)</span></span>


<span data-ttu-id="addfc-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="addfc-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="addfc-p101">Un objet **Parameter** représente une valeur fournie à une requête. Le paramètre est associé à un objet **QueryDef** créé à partir d'une requête avec paramètres</span><span class="sxs-lookup"><span data-stu-id="addfc-p101">A **Parameter** object represents a value supplied to a query. The parameter is associated with a **QueryDef** object created from a parameter query.</span></span>

## <a name="remarks"></a><span data-ttu-id="addfc-106">Remarques</span><span class="sxs-lookup"><span data-stu-id="addfc-106">Remarks</span></span>

<span data-ttu-id="addfc-107">Grâce aux objets **Parameter**, vous pouvez modifier les arguments d'un objet **QueryDef** exécuté fréquemment sans devoir recompiler la requête.</span><span class="sxs-lookup"><span data-stu-id="addfc-107">**Parameter** objects allow you to change the arguments in a frequently run **QueryDef** object without having to recompile the query.</span></span>

<span data-ttu-id="addfc-p102">À l'aide des propriétés d'un objet **Parameter**, vous pouvez définir un paramètre de requête pouvant être modifié avant l'exécution de la requête. Vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="addfc-p102">Using the properties of a **Parameter** object, you can set a query parameter that can be changed before the query is run. You can:</span></span>

  - <span data-ttu-id="addfc-110">Utiliser la propriété **Name** pour renvoyer le nom d'un paramètre.</span><span class="sxs-lookup"><span data-stu-id="addfc-110">Use the **Name** property to return the name of a parameter.</span></span>

  - <span data-ttu-id="addfc-111">Utiliser la propriété **Value** pour définir ou renvoyer les valeurs de paramètre à utiliser dans la requête.</span><span class="sxs-lookup"><span data-stu-id="addfc-111">Use the **Value** property to set or return the parameter values to be used in the query.</span></span>

  - <span data-ttu-id="addfc-112">Utiliser la propriété **Type** pour renvoyer le type de données de l'objet **Parameter**.</span><span class="sxs-lookup"><span data-stu-id="addfc-112">Use the **Type** property to return the data type of the **Parameter** object.</span></span>

  - <span data-ttu-id="addfc-113">Utiliser la propriété **Direction** pour indiquer si le paramètre est un paramètre d'entrée et/ou de sortie.</span><span class="sxs-lookup"><span data-stu-id="addfc-113">Use the **Direction** property to set or return whether the parameter is an input parameter, an output parameter, or both.</span></span>

## <a name="example"></a><span data-ttu-id="addfc-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="addfc-114">Example</span></span>

<span data-ttu-id="addfc-p103">Cet exemple démontre les objets **Parameter** et la collection **Parameters** en créant un objet **QueryDef** temporaire et en extrayant les données sur la base des modifications apportées aux objets **Parameters** de l'objet **QueryDef**. La procédure ParametersChange est requise pour exécuter cette opération.</span><span class="sxs-lookup"><span data-stu-id="addfc-p103">This example demonstrates **Parameter** objects and the **Parameters** collection by creating a temporary **QueryDef** and retrieving data based on changes made to the **QueryDef** object's **Parameters**. The ParametersChange procedure is required for this procedure to run.</span></span>

```vb
    Sub ParameterX() 
     
     Dim dbsNorthwind As Database 
     Dim qdfReport As QueryDef 
     Dim prmBegin As Parameter 
     Dim prmEnd As Parameter 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     
     ' Create temporary QueryDef object with two 
     ' parameters. 
     Set qdfReport = dbsNorthwind.CreateQueryDef("", _ 
     "PARAMETERS dteBegin DateTime, dteEnd DateTime; " & _ 
     "SELECT EmployeeID, COUNT(OrderID) AS NumOrders " & _ 
     "FROM Orders WHERE ShippedDate BETWEEN " & _ 
     "[dteBegin] AND [dteEnd] GROUP BY EmployeeID " & _ 
     "ORDER BY EmployeeID") 
     Set prmBegin = qdfReport.Parameters!dteBegin 
     Set prmEnd = qdfReport.Parameters!dteEnd 
     
     ' Print report using specified parameter values. 
     ParametersChange qdfReport, prmBegin, #1/1/95#, _ 
     prmEnd, #6/30/95# 
     ParametersChange qdfReport, prmBegin, #7/1/95#, _ 
     prmEnd, #12/31/95# 
     
     dbsNorthwind.Close 
     
    End Sub 
     
    Sub ParametersChange(qdfTemp As QueryDef, _ 
     prmFirst As Parameter, dteFirst As Date, _ 
     prmLast As Parameter, dteLast As Date) 
     ' Report function for ParameterX. 
     
     Dim rstTemp As Recordset 
     Dim fldLoop As Field 
     
     ' Set parameter values and open recordset from 
     ' temporary QueryDef object. 
     prmFirst = dteFirst 
     prmLast = dteLast 
     Set rstTemp = _ 
     qdfTemp.OpenRecordset(dbOpenForwardOnly) 
     Debug.Print "Period " & dteFirst & " to " & dteLast 
     
     ' Enumerate recordset. 
     Do While Not rstTemp.EOF 
     
     ' Enumerate Fields collection of recordset. 
     For Each fldLoop In rstTemp.Fields 
     Debug.Print " - " & fldLoop.Name & " = " & fldLoop; 
     Next fldLoop 
     
     Debug.Print 
     rstTemp.MoveNext 
     Loop 
     
     rstTemp.Close 
     
    End Sub
```