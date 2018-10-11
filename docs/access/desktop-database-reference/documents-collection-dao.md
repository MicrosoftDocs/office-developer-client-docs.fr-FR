---
title: Documents Collection (DAO)
TOCTitle: Documents Collection
ms:assetid: ae2fef58-34e7-eea6-ca51-d3903432c7f5
ms:mtpsurl: https://msdn.microsoft.com/library/Ff821742(v=office.15)
ms:contentKeyID: 48547063
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 92a9dc96b438a55401df95a63ea47b85d96b0f83
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472401"
---
# <a name="documents-collection-dao"></a><span data-ttu-id="dfc45-102">Documents Collection (DAO)</span><span class="sxs-lookup"><span data-stu-id="dfc45-102">Documents Collection (DAO)</span></span>


<span data-ttu-id="dfc45-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="dfc45-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="dfc45-104">Une collection **Documents** contient tous les objets **Document** d'un type spécifique d'objet (bases de données de moteur de base de données Microsoft Access uniquement).</span><span class="sxs-lookup"><span data-stu-id="dfc45-104">A **Documents** collection contains all of the **Document** objects for a specific type of object (Microsoft Access database engine databases only).</span></span>

## <a name="remarks"></a><span data-ttu-id="dfc45-105">Remarques</span><span class="sxs-lookup"><span data-stu-id="dfc45-105">Remarks</span></span>

<span data-ttu-id="dfc45-106">Chaque objet **Container** est doté d'une collection **Documents** contenant les objets **Document** qui décrivent les instances des objets intégrés du type spécifié dans l'objet **Container**.</span><span class="sxs-lookup"><span data-stu-id="dfc45-106">Each **Container** object has a **Documents** collection containing **Document** objects that describe instances of built-in objects of the type specified by the **Container**.</span></span>

<span data-ttu-id="dfc45-107">Pour faire référence à un objet **Document** d'une collection par son numéro ordinal ou par son paramètre de propriété **Name**, utilisez l'une des syntaxes suivantes :</span><span class="sxs-lookup"><span data-stu-id="dfc45-107">To refer to a **Document** object in a collection by its ordinal number or by its **Name** property setting, use any of the following syntax forms:</span></span>

  - <span data-ttu-id="dfc45-108">**Documents**(0)</span><span class="sxs-lookup"><span data-stu-id="dfc45-108">**Documents**(0)</span></span>

  - <span data-ttu-id="dfc45-109">**Documents** («*nom*»)</span><span class="sxs-lookup"><span data-stu-id="dfc45-109">**Documents**("*name*")</span></span>

  - <span data-ttu-id="dfc45-110">**Documents**\!\[*nom*\]</span><span class="sxs-lookup"><span data-stu-id="dfc45-110">**Documents**\!\[*name*\]</span></span>

## <a name="example"></a><span data-ttu-id="dfc45-111">Exemple</span><span class="sxs-lookup"><span data-stu-id="dfc45-111">Example</span></span>

<span data-ttu-id="dfc45-112">Cet exemple énumère la collection **Documents** du conteneur Tables, puis la collection **Properties** du premier objet **Document** de la collection.</span><span class="sxs-lookup"><span data-stu-id="dfc45-112">This example enumerates the **Documents** collection of the Tables container, and then enumerates the **Properties** collection of the first **Document** object in the collection.</span></span>

```vb 
Sub DocumentX() 
 
 Dim dbsNorthwind As Database 
 Dim docLoop As Document 
 Dim prpLoop As Property 
 
 Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
 With dbsNorthwind.Containers!Tables 
 Debug.Print "Documents in " & .Name & " container" 
 ' Enumerate the Documents collection of the Tables 
 ' container. 
 For Each docLoop In .Documents 
 Debug.Print " " & docLoop.Name 
 Next docLoop 
 With .Documents(0) 
 ' Enumerate the Properties collection of the first. 
 ' Document object of the Tables container. 
 Debug.Print "Properties of " & .Name & " document" 
 On Error Resume Next 
 For Each prpLoop In .Properties 
 Debug.Print " " & prpLoop.Name & " = " & _ 
 prpLoop 
 Next prpLoop 
 On Error GoTo 0 
 End With 
 End With 
 
 dbsNorthwind.Close 
 
End Sub 
 
```
