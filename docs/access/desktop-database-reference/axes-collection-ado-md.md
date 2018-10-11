---
title: Axes, collection (ADO MD)
TOCTitle: Axes Collection (ADO MD)
ms:assetid: 7c719197-45f1-a5b9-665d-25cb693b1eb0
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249520(v=office.15)
ms:contentKeyID: 48545836
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: d2b078f2d8f1ea6562d6f82f90943923c7d92a07
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471620"
---
# <a name="axes-collection-ado-md"></a><span data-ttu-id="56c65-102">Axes, collection (ADO MD)</span><span class="sxs-lookup"><span data-stu-id="56c65-102">Axes Collection (ADO MD)</span></span>


<span data-ttu-id="56c65-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="56c65-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="56c65-104">Contient les objets [Axis](axis-object-ado-md.md) qui définissent un ensemble de cellules.</span><span class="sxs-lookup"><span data-stu-id="56c65-104">Contains the [Axis](axis-object-ado-md.md) objects that define a cellset.</span></span>

## <a name="remarks"></a><span data-ttu-id="56c65-105">Remarques</span><span class="sxs-lookup"><span data-stu-id="56c65-105">Remarks</span></span>

<span data-ttu-id="56c65-p101">Un objet [Cellset](cellset-object-ado-md.md) contient une collection **Axes**. Une fois l' **ensemble de cellules** ouvert, cette collection renferme au moins un **axe**. Reportez-vous à la rubrique consacrée à l'objet [Axis](axis-object-ado-md.md) pour une explication plus détaillée de l'utilisation des objets **Axis**.</span><span class="sxs-lookup"><span data-stu-id="56c65-p101">A [Cellset](cellset-object-ado-md.md) object contains an **Axes** collection. Once the **Cellset** is opened, this collection will contain at least one **Axis**. See the [Axis](axis-object-ado-md.md) object for a more detailed explanation of how to use **Axis** objects.</span></span>


> [!NOTE]
> <P><span data-ttu-id="56c65-p102">[!REMARQUE] L'axe de filtrage d'un <STRONG>ensemble de cellules</STRONG> ne figure pas dans la collection <STRONG>Axes</STRONG>. Reportez-vous à la rubrique consacrée à la propriété <A href="filteraxis-property-ado-md.md">FilterAxis</A> pour plus d'informations.</span><span class="sxs-lookup"><span data-stu-id="56c65-p102">The filter axis of a <STRONG>Cellset</STRONG> is not contained in the <STRONG>Axes</STRONG> collection. See the <A href="filteraxis-property-ado-md.md">FilterAxis</A> property for more information.</span></span></P>



<span data-ttu-id="56c65-p103">**Axes** est une collection ADO standard. Avec les propriétés et méthodes d'une collection, vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="56c65-p103">**Axes** is a standard ADO collection. With the properties and methods of a collection, you can do the following:</span></span>

  - <span data-ttu-id="56c65-113">Obtenir le nombre d'objets d'une collection à l'aide de la propriété [Count](count-property-ado.md).</span><span class="sxs-lookup"><span data-stu-id="56c65-113">Obtain the number of objects in the collection with the [Count](count-property-ado.md) property.</span></span>

  - <span data-ttu-id="56c65-114">Renvoyer un objet de la collection à l'aide de la propriété [Item](item-property-ado.md) par défaut.</span><span class="sxs-lookup"><span data-stu-id="56c65-114">Return an object from the collection with the default [Item](item-property-ado.md) property.</span></span>

  - <span data-ttu-id="56c65-115">Mettre à jour les objets de la collection à partir du fournisseur à l'aide de la méthode [Refresh](refresh-method-ado.md).</span><span class="sxs-lookup"><span data-stu-id="56c65-115">Update the objects in the collection from the provider with the [Refresh](refresh-method-ado.md) method.</span></span>
