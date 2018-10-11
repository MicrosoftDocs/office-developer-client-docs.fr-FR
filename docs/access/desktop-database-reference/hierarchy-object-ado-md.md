---
title: Hierarchy, objet (ADO MD)
TOCTitle: Hierarchy Object (ADO MD)
ms:assetid: 26e4e690-59ad-fb87-66b0-f3310df42d0c
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249031(v=office.15)
ms:contentKeyID: 48543825
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 34513e3188653bdba04add376f06911ce86386bd
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472000"
---
# <a name="hierarchy-object-ado-md"></a><span data-ttu-id="660b2-102">Hierarchy, objet (ADO MD)</span><span class="sxs-lookup"><span data-stu-id="660b2-102">Hierarchy Object (ADO MD)</span></span>


<span data-ttu-id="660b2-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="660b2-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="660b2-p101">Représente un moyen d'agréger ou de « cumuler » les membres d'une [dimension](dimension-object-ado-md.md). Vous pouvez agréger une dimension le long d'une ou de plusieurs hiérarchies.</span><span class="sxs-lookup"><span data-stu-id="660b2-p101">Represents one way in which the members of a [dimension](dimension-object-ado-md.md) can be aggregated or "rolled up." A dimension can be aggregated along one or more hierarchies.</span></span>

## <a name="remarks"></a><span data-ttu-id="660b2-106">Remarques</span><span class="sxs-lookup"><span data-stu-id="660b2-106">Remarks</span></span>

<span data-ttu-id="660b2-107">Avec les collections et propriétés d'un objet **Hierarchy**, vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="660b2-107">With the collections and properties of a **Hierarchy** object, you can do the following:</span></span>

  - <span data-ttu-id="660b2-108">Identifier la **hiérarchie** à l'aide des propriétés [Name](name-property-ado-md.md) et [UniqueName](uniquename-property-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="660b2-108">Identify the **Hierarchy** with the [Name](name-property-ado-md.md) and [UniqueName](uniquename-property-ado-md.md) properties.</span></span>

  - <span data-ttu-id="660b2-109">Renvoyer une chaîne significative qui décrit la **hiérarchie** à l'aide de la propriété [Description](description-property-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="660b2-109">Return a meaningful string that describes the **Hierarchy** with the [Description](description-property-ado-md.md) property.</span></span>

  - <span data-ttu-id="660b2-110">Renvoyer les objets [Level](level-object-ado-md.md) qui constituent la **hiérarchie** à l'aide de la collection [Levels](levels-collection-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="660b2-110">Return the [Level](level-object-ado-md.md) objects that make up the **Hierarchy** with the [Levels](levels-collection-ado-md.md) collection.</span></span>

  - <span data-ttu-id="660b2-111">Utiliser la collection ADO standard [Properties](properties-collection-ado.md) pour obtenir des informations supplémentaires à propos de l'objet **Hierarchy**.</span><span class="sxs-lookup"><span data-stu-id="660b2-111">Use the standard ADO [Properties](properties-collection-ado.md) collection to obtain additional information about the **Hierarchy** object.</span></span>

<span data-ttu-id="660b2-p102">La collection **Properties** renferme les propriétés fournies par le fournisseur. Le tableau suivant dresse la liste des propriétés potentiellement disponibles. La liste réelle des propriétés peut varier en fonction de la mise en œuvre du fournisseur. Reportez-vous à la documentation de votre fournisseur pour une liste plus complète des propriétés disponibles.</span><span class="sxs-lookup"><span data-stu-id="660b2-p102">The **Properties** collection contains provider-supplied properties. The following table lists properties that might be available. The actual property list may differ depending upon the implementation of the provider. See the documentation for your provider for a more complete list of available properties.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="660b2-116">Nom</span><span class="sxs-lookup"><span data-stu-id="660b2-116">Name</span></span></p></th>
<th><p><span data-ttu-id="660b2-117">Description</span><span class="sxs-lookup"><span data-stu-id="660b2-117">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="660b2-118">AllMember</span><span class="sxs-lookup"><span data-stu-id="660b2-118">AllMember</span></span></p></td>
<td><p><span data-ttu-id="660b2-119">Le membre au niveau le plus élevé de la hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="660b2-119">The member at the highest level of rollup in the hierarchy.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="660b2-120">Nom de catalogue</span><span class="sxs-lookup"><span data-stu-id="660b2-120">CatalogName</span></span></p></td>
<td><p><span data-ttu-id="660b2-121">Le nom du catalogue auquel ce cube appartient.</span><span class="sxs-lookup"><span data-stu-id="660b2-121">The name of the catalog to which this cube belongs.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="660b2-122">Nom du cube</span><span class="sxs-lookup"><span data-stu-id="660b2-122">CubeName</span></span></p></td>
<td><p><span data-ttu-id="660b2-123">Le nom du cube.</span><span class="sxs-lookup"><span data-stu-id="660b2-123">The name of the cube.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="660b2-124">DefaultMember</span><span class="sxs-lookup"><span data-stu-id="660b2-124">DefaultMember</span></span></p></td>
<td><p><span data-ttu-id="660b2-125">Le nom unique du membre par défaut de cette hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="660b2-125">The unique name of the default member for this hierarchy.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="660b2-126">Description</span><span class="sxs-lookup"><span data-stu-id="660b2-126">Description</span></span></p></td>
<td><p><span data-ttu-id="660b2-127">Une description significative de la hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="660b2-127">A meaningful description of the hierarchy.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="660b2-128">DimensionType</span><span class="sxs-lookup"><span data-stu-id="660b2-128">DimensionType</span></span></p></td>
<td><p><span data-ttu-id="660b2-129">Le type de dimension à laquelle cette hiérarchie appartient.</span><span class="sxs-lookup"><span data-stu-id="660b2-129">The type of dimension to which this hierarchy belongs.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="660b2-130">DimensionUniqueName</span><span class="sxs-lookup"><span data-stu-id="660b2-130">DimensionUniqueName</span></span></p></td>
<td><p><span data-ttu-id="660b2-131">Le nom non ambigu de la dimension.</span><span class="sxs-lookup"><span data-stu-id="660b2-131">The unambiguous name of the dimension.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="660b2-132">HierarchyCaption</span><span class="sxs-lookup"><span data-stu-id="660b2-132">HierarchyCaption</span></span></p></td>
<td><p><span data-ttu-id="660b2-133">Une étiquette ou une légende associée à la hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="660b2-133">A label or caption associated with the hierarchy.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="660b2-134">HierarchyCardinality</span><span class="sxs-lookup"><span data-stu-id="660b2-134">HierarchyCardinality</span></span></p></td>
<td><p><span data-ttu-id="660b2-135">Le nombre de membres de la hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="660b2-135">The number of members in the hierarchy.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="660b2-136">HierarchyGUID</span><span class="sxs-lookup"><span data-stu-id="660b2-136">HierarchyGUID</span></span></p></td>
<td><p><span data-ttu-id="660b2-137">Le GUID de la hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="660b2-137">The GUID of the hierarchy.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="660b2-138">HierarchyName</span><span class="sxs-lookup"><span data-stu-id="660b2-138">HierarchyName</span></span></p></td>
<td><p><span data-ttu-id="660b2-139">Le nom de la hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="660b2-139">The name of the hierarchy.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="660b2-140">HierarchyUniqueName</span><span class="sxs-lookup"><span data-stu-id="660b2-140">HierarchyUniqueName</span></span></p></td>
<td><p><span data-ttu-id="660b2-141">Le nom non ambigu de la hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="660b2-141">The unambiguous name of the hierarchy.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="660b2-142">SchemaName</span><span class="sxs-lookup"><span data-stu-id="660b2-142">SchemaName</span></span></p></td>
<td><p><span data-ttu-id="660b2-143">Le nom du schéma auquel ce cube appartient.</span><span class="sxs-lookup"><span data-stu-id="660b2-143">The name of the schema to which this cube belongs.</span></span></p></td>
</tr>
</tbody>
</table>
