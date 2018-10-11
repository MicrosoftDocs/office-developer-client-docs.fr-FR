---
title: ChildCount, propriété (ADO MD)
TOCTitle: ChildCount Property (ADO MD)
ms:assetid: ff1872f0-d5f6-174e-0473-7997a462ca81
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250309(v=office.15)
ms:contentKeyID: 48548956
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 3fb6edf995651241d15b9bc9fc9df1ce507801bc
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470710"
---
# <a name="childcount-property-ado-md"></a><span data-ttu-id="f699c-102">ChildCount, propriété (ADO MD)</span><span class="sxs-lookup"><span data-stu-id="f699c-102">ChildCount Property (ADO MD)</span></span>


<span data-ttu-id="f699c-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="f699c-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="f699c-104">Indique le nombre de membres dont l'objet [Member](member-object-ado-md.md) actif est le parent dans une hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="f699c-104">Indicates the number of members for which the current [Member](member-object-ado-md.md) object is the parent in a hierarchy.</span></span>

## <a name="return-values"></a><span data-ttu-id="f699c-105">Valeurs de retour</span><span class="sxs-lookup"><span data-stu-id="f699c-105">Return Values</span></span>

<span data-ttu-id="f699c-106">Renvoie un nombre entier de type **Long** et est en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="f699c-106">Returns a **Long** integer and is read-only.</span></span>

## <a name="remarks"></a><span data-ttu-id="f699c-107">Notes</span><span class="sxs-lookup"><span data-stu-id="f699c-107">Remarks</span></span>

<span data-ttu-id="f699c-p101">Utilisez la propriété **ChildCount** pour retourner une estimation du nombre d'enfants d'un objet **Member**. La propriété **Children** peut renvoyer le nombre réel d'enfants d'un objet [Member](children-property-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="f699c-p101">Use the **ChildCount** property to return an estimate of how many children a **Member** has. The actual children of a **Member** can be returned by the [Children](children-property-ado-md.md) property.</span></span>

<span data-ttu-id="f699c-p102">Pour les objets **Member** d'un objet [Position](position-object-ado-md.md), le nombre maximum renvoyé est 65 536. Si le nombre réel d'enfants dépasse 65 536, la valeur de retour sera toujours 65 536. L'application doit normalement interpréter une propriété **ChildCount** d'une valeur de 65 536 comme égale ou supérieure à 65 536 enfants.</span><span class="sxs-lookup"><span data-stu-id="f699c-p102">For **Member** objects from a [Position](position-object-ado-md.md) object, the maximum number returned is 65536. If the actual number of children exceeds 65536, the value returned will still be 65536. Therefore, the application should interpret a **ChildCount** of 65536 as equal to or greater than 65536 children.</span></span>

<span data-ttu-id="f699c-p103">Pour les objets **Member** d'un objet [Level](level-object-ado-md.md), utilisez la propriété [Count](count-property-ado.md) de la collection ADO sur la collection **Children** afin de déterminer le nombre exact d'enfants. Ce calcul peut prendre du temps si la collection contient un grand nombre d'enfants.</span><span class="sxs-lookup"><span data-stu-id="f699c-p103">For **Member** objects from a [Level](level-object-ado-md.md) object, use the ADO collection [Count](count-property-ado.md) property on the **Children** collection to determine the exact number of children. Determining the exact number of children may be slow if the number of children in the collection is large.</span></span>
