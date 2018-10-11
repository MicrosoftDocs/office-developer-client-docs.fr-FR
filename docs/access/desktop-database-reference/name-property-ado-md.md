---
title: Name, propriété (ADO MD)
TOCTitle: Name Property (ADO MD)
ms:assetid: 31ea6dad-c464-3af7-4b7a-086900656c2c
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249093(v=office.15)
ms:contentKeyID: 48544065
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 0cc7a89e0d6b9cdaed2c54d3269b61280ce3306c
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469877"
---
# <a name="name-property-ado-md"></a><span data-ttu-id="c4a4b-102">Name, propriété (ADO MD)</span><span class="sxs-lookup"><span data-stu-id="c4a4b-102">Name Property (ADO MD)</span></span>


<span data-ttu-id="c4a4b-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="c4a4b-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="c4a4b-104">Indique le nom d'un objet.</span><span class="sxs-lookup"><span data-stu-id="c4a4b-104">Indicates the name of an object.</span></span>

## <a name="return-values"></a><span data-ttu-id="c4a4b-105">Valeurs de retour</span><span class="sxs-lookup"><span data-stu-id="c4a4b-105">Return Values</span></span>

<span data-ttu-id="c4a4b-106">Retourne une valeur de type **String** et est en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="c4a4b-106">Returns a **String** and is read-only.</span></span>

## <a name="remarks"></a><span data-ttu-id="c4a4b-107">Notes</span><span class="sxs-lookup"><span data-stu-id="c4a4b-107">Remarks</span></span>

<span data-ttu-id="c4a4b-p101">Vous pouvez récupérer la propriété **Name** d'un objet par référence ordinale, après quoi vous pouvez directement faire référence à l'objet par son nom. Si, par exemple, cdf.CubeDefs(0).Name produit la valeur « Bobs Video Store », vous pouvez faire référence à cet objet [CubeDef](cubedef-object-ado-md.md) comme suit : cdf.CubeDefs(« Bobs Video Store »).</span><span class="sxs-lookup"><span data-stu-id="c4a4b-p101">You can retrieve the **Name** property of an object by an ordinal reference, after which you can refer to the object directly by name. For example, if cdf.CubeDefs(0).Name yields "Bobs Video Store", you can refer to this [CubeDef](cubedef-object-ado-md.md) as cdf.CubeDefs("Bobs Video Store").</span></span>
