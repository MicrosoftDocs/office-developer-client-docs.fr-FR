---
title: FilterAxis, propriété (ADO MD)
TOCTitle: FilterAxis Property (ADO MD)
ms:assetid: 36720d77-4b16-1d17-6d80-d35265f4a8ad
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249124(v=office.15)
ms:contentKeyID: 48544172
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 912dfec595f718c2144e69b4fbd3af592f8b6d5b
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470302"
---
# <a name="filteraxis-property-ado-md"></a><span data-ttu-id="e14a6-102">FilterAxis, propriété (ADO MD)</span><span class="sxs-lookup"><span data-stu-id="e14a6-102">FilterAxis Property (ADO MD)</span></span>


<span data-ttu-id="e14a6-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="e14a6-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="e14a6-104">Fournit des informations de filtrage concernant l'ensemble de cellules actif.</span><span class="sxs-lookup"><span data-stu-id="e14a6-104">Indicates filter information about the current cellset.</span></span>

## <a name="return-values"></a><span data-ttu-id="e14a6-105">Valeurs de retour</span><span class="sxs-lookup"><span data-stu-id="e14a6-105">Return Values</span></span>

<span data-ttu-id="e14a6-106">Retourne un objet [Axis](axis-object-ado-md.md) et est en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="e14a6-106">Returns an [Axis](axis-object-ado-md.md) object, and is read-only.</span></span>

## <a name="remarks"></a><span data-ttu-id="e14a6-107">Notes</span><span class="sxs-lookup"><span data-stu-id="e14a6-107">Remarks</span></span>

<span data-ttu-id="e14a6-p101">Utilisez la propriété **FilterAxis** pour retourner les informations sur les dimensions utilisées pour découper les données. La propriété [DimensionCount](dimensioncount-property-ado-md.md) de l'objet **Axis** retourne le nombre de dimensions de découpage. Cet axe comprend généralement une seule ligne.</span><span class="sxs-lookup"><span data-stu-id="e14a6-p101">Use the **FilterAxis** property to return information about the dimensions that were used to slice the data. The [DimensionCount](dimensioncount-property-ado-md.md) property of the **Axis** returns the number of slicer dimensions. This axis usually has just one row.</span></span>

<span data-ttu-id="e14a6-111">L'objet **Axis** retourné par [FilterAxis](filteraxis-property-ado-md.md) n'est pas compris dans la collection [Axes](axes-collection-ado-md.md) d'un objet [Cellset](cellset-object-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="e14a6-111">The **Axis** returned by [FilterAxis](filteraxis-property-ado-md.md) is not contained in the [Axes](axes-collection-ado-md.md) collection for a [Cellset](cellset-object-ado-md.md) object.</span></span>
