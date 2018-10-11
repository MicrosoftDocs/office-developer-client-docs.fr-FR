---
title: Dimension, objet (ADO MD)
TOCTitle: Dimension Object (ADO MD)
ms:assetid: 12f43cfc-c74e-a2e8-7f6e-75fc68472c4b
ms:mtpsurl: https://msdn.microsoft.com/library/JJ248902(v=office.15)
ms:contentKeyID: 48543355
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 15ae89905c9ca82c4c41701c4ed713c02722e322
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469132"
---
# <a name="dimension-object-ado-md"></a><span data-ttu-id="5cc67-102">Dimension, objet (ADO MD)</span><span class="sxs-lookup"><span data-stu-id="5cc67-102">Dimension Object (ADO MD)</span></span>


<span data-ttu-id="5cc67-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="5cc67-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="5cc67-104">Représente une des dimensions d'un cube multidimensionnel, contenant une ou plusieurs hiérarchies de membres.</span><span class="sxs-lookup"><span data-stu-id="5cc67-104">Represents one of the dimensions of a multidimensional cube, containing one or more hierarchies of members.</span></span>

## <a name="remarks"></a><span data-ttu-id="5cc67-105">Remarques</span><span class="sxs-lookup"><span data-stu-id="5cc67-105">Remarks</span></span>

<span data-ttu-id="5cc67-106">Avec les collections et propriétés d'un objet **Dimension**, vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="5cc67-106">With the collections and properties of a **Dimension** object, you can do the following:</span></span>

  - <span data-ttu-id="5cc67-107">Identifier la **dimension** à l'aide des propriétés [Name](name-property-ado-md.md) et [UniqueName](uniquename-property-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="5cc67-107">Identify the **Dimension** with the [Name](name-property-ado-md.md) and [UniqueName](uniquename-property-ado-md.md) properties.</span></span>

  - <span data-ttu-id="5cc67-108">Renvoyer une chaîne significative qui décrit la **dimension** à l'aide de la propriété [Description](description-property-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="5cc67-108">Return a meaningful string that describes the **Dimension** with the [Description](description-property-ado-md.md) property.</span></span>

  - <span data-ttu-id="5cc67-109">Renvoyer les objets [Hierarchy](hierarchy-object-ado-md.md) qui constituent la **dimension** à l'aide de la collection [Hierarchies](hierarchies-collection-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="5cc67-109">Return the [Hierarchy](hierarchy-object-ado-md.md) objects that make up the **Dimension** with the [Hierarchies](hierarchies-collection-ado-md.md) collection.</span></span>

  - <span data-ttu-id="5cc67-110">Utiliser la collection ADO standard [Properties](properties-collection-ado.md) pour obtenir des informations supplémentaires à propos de l'objet **Dimension**.</span><span class="sxs-lookup"><span data-stu-id="5cc67-110">Use the standard ADO [Properties](properties-collection-ado.md) collection to obtain additional information about the **Dimension** object.</span></span>

<span data-ttu-id="5cc67-p101">La collection **Properties** renferme les propriétés fournies par le fournisseur. Le tableau suivant dresse la liste des propriétés potentiellement disponibles. La liste réelle des propriétés peut varier en fonction de la mise en œuvre du fournisseur. Reportez-vous à la documentation de votre fournisseur pour une liste plus complète des propriétés disponibles.</span><span class="sxs-lookup"><span data-stu-id="5cc67-p101">The **Properties** collection contains provider-supplied properties. The following table lists properties that might be available. The actual property list may differ depending upon the implementation of the provider. See the documentation for your provider for a more complete list of available properties.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="5cc67-115">Nom</span><span class="sxs-lookup"><span data-stu-id="5cc67-115">Name</span></span></p></th>
<th><p><span data-ttu-id="5cc67-116">Description</span><span class="sxs-lookup"><span data-stu-id="5cc67-116">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="5cc67-117">Nom de catalogue</span><span class="sxs-lookup"><span data-stu-id="5cc67-117">CatalogName</span></span></p></td>
<td><p><span data-ttu-id="5cc67-118">Le nom du catalogue auquel ce cube appartient.</span><span class="sxs-lookup"><span data-stu-id="5cc67-118">The name of the catalog to which this cube belongs.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="5cc67-119">Nom du cube</span><span class="sxs-lookup"><span data-stu-id="5cc67-119">CubeName</span></span></p></td>
<td><p><span data-ttu-id="5cc67-120">Le nom du cube.</span><span class="sxs-lookup"><span data-stu-id="5cc67-120">The name of the cube.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="5cc67-121">DefaultHierarchy</span><span class="sxs-lookup"><span data-stu-id="5cc67-121">DefaultHierarchy</span></span></p></td>
<td><p><span data-ttu-id="5cc67-122">Le nom unique de la hiérarchie par défaut.</span><span class="sxs-lookup"><span data-stu-id="5cc67-122">The unique name of the default hierarchy.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="5cc67-123">Description</span><span class="sxs-lookup"><span data-stu-id="5cc67-123">Description</span></span></p></td>
<td><p><span data-ttu-id="5cc67-124">Une description significative du cube.</span><span class="sxs-lookup"><span data-stu-id="5cc67-124">A meaningful description of the cube.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="5cc67-125">DimensionCaption</span><span class="sxs-lookup"><span data-stu-id="5cc67-125">DimensionCaption</span></span></p></td>
<td><p><span data-ttu-id="5cc67-126">Une étiquette ou une légende associée à la dimension.</span><span class="sxs-lookup"><span data-stu-id="5cc67-126">A label or caption associated with the dimension.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="5cc67-127">DimensionCardinality</span><span class="sxs-lookup"><span data-stu-id="5cc67-127">DimensionCardinality</span></span></p></td>
<td><p><span data-ttu-id="5cc67-128">Le nombre de membres de la dimension.</span><span class="sxs-lookup"><span data-stu-id="5cc67-128">The number of members in the dimension.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="5cc67-129">DimensionGUID</span><span class="sxs-lookup"><span data-stu-id="5cc67-129">DimensionGUID</span></span></p></td>
<td><p><span data-ttu-id="5cc67-130">Le GUID de la dimension.</span><span class="sxs-lookup"><span data-stu-id="5cc67-130">The GUID of the dimension.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="5cc67-131">Nomdimension</span><span class="sxs-lookup"><span data-stu-id="5cc67-131">DimensionName</span></span></p></td>
<td><p><span data-ttu-id="5cc67-132">Le nom de la dimension.</span><span class="sxs-lookup"><span data-stu-id="5cc67-132">The name of the dimension.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="5cc67-133">DimensionOrdinal</span><span class="sxs-lookup"><span data-stu-id="5cc67-133">DimensionOrdinal</span></span></p></td>
<td><p><span data-ttu-id="5cc67-134">Le numéro ordinal de la dimension parmi le groupe de dimensions constituant le cube.</span><span class="sxs-lookup"><span data-stu-id="5cc67-134">The ordinal number of the dimension among the group of dimensions that form the cube.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="5cc67-135">DimensionType</span><span class="sxs-lookup"><span data-stu-id="5cc67-135">DimensionType</span></span></p></td>
<td><p><span data-ttu-id="5cc67-136">Le type de dimension.</span><span class="sxs-lookup"><span data-stu-id="5cc67-136">The dimension type.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="5cc67-137">DimensionUniqueName</span><span class="sxs-lookup"><span data-stu-id="5cc67-137">DimensionUniqueName</span></span></p></td>
<td><p><span data-ttu-id="5cc67-138">Le nom non ambigu de la dimension.</span><span class="sxs-lookup"><span data-stu-id="5cc67-138">The unambiguous name of the dimension.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="5cc67-139">SchemaName</span><span class="sxs-lookup"><span data-stu-id="5cc67-139">SchemaName</span></span></p></td>
<td><p><span data-ttu-id="5cc67-140">Le nom du schéma auquel ce cube appartient.</span><span class="sxs-lookup"><span data-stu-id="5cc67-140">The name of the schema to which this cube belongs.</span></span></p></td>
</tr>
</tbody>
</table>
