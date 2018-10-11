---
title: Cell, objet (ADO MD)
TOCTitle: Cell Object (ADO MD)
ms:assetid: b9d00b71-1f40-5bd1-4b89-fbdb59c552ba
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249892(v=office.15)
ms:contentKeyID: 48547356
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: f38de1adc48e7d63a3b67e45f242a8cc884633ac
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471579"
---
# <a name="cell-object-ado-md"></a><span data-ttu-id="04c10-102">Cell, objet (ADO MD)</span><span class="sxs-lookup"><span data-stu-id="04c10-102">Cell Object (ADO MD)</span></span>


<span data-ttu-id="04c10-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="04c10-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="04c10-104">Représente les données à l'intersection des coordonnées des axes d'un ensemble de cellules.</span><span class="sxs-lookup"><span data-stu-id="04c10-104">Represents the data at the intersection of axis coordinates contained in a cellset.</span></span>

## <a name="remarks"></a><span data-ttu-id="04c10-105">Remarques</span><span class="sxs-lookup"><span data-stu-id="04c10-105">Remarks</span></span>

<span data-ttu-id="04c10-106">Un objet **Cell** est renvoyé par la propriété [Item](item-property-ado-md-cellset.md) d'un objet [Cellset](cellset-object-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="04c10-106">A **Cell** object is returned by the [Item](item-property-ado-md-cellset.md) property of a [Cellset](cellset-object-ado-md.md) object.</span></span>

<span data-ttu-id="04c10-107">Avec les collections et propriétés d'un objet **Cell**, vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="04c10-107">With the collections and properties of a **Cell** object, you can do the following:</span></span>

  - <span data-ttu-id="04c10-108">Renvoyer les données de la **cellule** à l'aide de la propriété [Value](value-property-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="04c10-108">Return the data in the **Cell** with the [Value](value-property-ado-md.md) property.</span></span>

  - <span data-ttu-id="04c10-109">Renvoyer la chaîne représentant l'affichage mis en forme de la propriété **Value** à l'aide de la propriété [FormattedValue](formattedvalue-property-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="04c10-109">Return the string representing the formatted display of the **Value** property with the [FormattedValue](formattedvalue-property-ado-md.md) property.</span></span>

  - <span data-ttu-id="04c10-110">Renvoyer la valeur ordinale de la **cellule** au sein de l' **ensemble de cellules** à l'aide de la propriété [Ordinal](ordinal-property-ado-md-cell.md).</span><span class="sxs-lookup"><span data-stu-id="04c10-110">Return the ordinal value of the **Cell** within the **Cellset** with the [Ordinal](ordinal-property-ado-md-cell.md) property.</span></span>

  - <span data-ttu-id="04c10-111">Déterminer la position de la **cellule** au sein de l'objet [CubeDef](cubedef-object-ado-md.md) à l'aide de la collection [Positions](positions-collection-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="04c10-111">Determine the position of the **Cell** within the [CubeDef](cubedef-object-ado-md.md) with the [Positions](positions-collection-ado-md.md) collection.</span></span>

  - <span data-ttu-id="04c10-112">Extraire d'autres informations relatives à la **cellule** à l'aide de la collection ADO [Properties](properties-collection-ado.md) standard.</span><span class="sxs-lookup"><span data-stu-id="04c10-112">Retrieve other information about the **Cell** with the standard ADO [Properties](properties-collection-ado.md) collection.</span></span>

<span data-ttu-id="04c10-p101">La collection **Properties** renferme les propriétés fournies par le fournisseur. Le tableau suivant dresse la liste des propriétés potentiellement disponibles. La liste réelle des propriétés peut varier en fonction de la mise en œuvre du fournisseur. Reportez-vous à la documentation de votre fournisseur pour une liste plus complète des propriétés disponibles.</span><span class="sxs-lookup"><span data-stu-id="04c10-p101">The **Properties** collection contains provider-supplied properties. The following table lists properties that might be available. The actual property list may differ depending upon the implementation of the provider. See the documentation for your provider for a more complete list of available properties.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="04c10-117">Nom</span><span class="sxs-lookup"><span data-stu-id="04c10-117">Name</span></span></p></th>
<th><p><span data-ttu-id="04c10-118">Description</span><span class="sxs-lookup"><span data-stu-id="04c10-118">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="04c10-119">BackColor</span><span class="sxs-lookup"><span data-stu-id="04c10-119">BackColor</span></span></p></td>
<td><p><span data-ttu-id="04c10-120">Couleur de fond utilisée pour l'affichage de la cellule.</span><span class="sxs-lookup"><span data-stu-id="04c10-120">Background color used when displaying the cell.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="04c10-121">FontFlags</span><span class="sxs-lookup"><span data-stu-id="04c10-121">FontFlags</span></span></p></td>
<td><p><span data-ttu-id="04c10-122">Masque binaire détaillant les effets de la police.</span><span class="sxs-lookup"><span data-stu-id="04c10-122">Bitmask detailing effects on the font.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="04c10-123">FontName</span><span class="sxs-lookup"><span data-stu-id="04c10-123">FontName</span></span></p></td>
<td><p><span data-ttu-id="04c10-124">Police utilisée pour afficher la valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="04c10-124">Font used to display the cell value.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="04c10-125">FontSize</span><span class="sxs-lookup"><span data-stu-id="04c10-125">FontSize</span></span></p></td>
<td><p><span data-ttu-id="04c10-126">Taille de police utilisée pour afficher la valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="04c10-126">Font size used to display the cell value.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="04c10-127">ForeColor</span><span class="sxs-lookup"><span data-stu-id="04c10-127">ForeColor</span></span></p></td>
<td><p><span data-ttu-id="04c10-128">Couleur de premier plan utilisée pour l'affichage de la cellule.</span><span class="sxs-lookup"><span data-stu-id="04c10-128">Foreground color used when displaying the cell.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="04c10-129">FormatString</span><span class="sxs-lookup"><span data-stu-id="04c10-129">FormatString</span></span></p></td>
<td><p><span data-ttu-id="04c10-130">Valeur dans une chaîne mise en forme.</span><span class="sxs-lookup"><span data-stu-id="04c10-130">Value in a formatted string.</span></span></p></td>
</tr>
</tbody>
</table>
