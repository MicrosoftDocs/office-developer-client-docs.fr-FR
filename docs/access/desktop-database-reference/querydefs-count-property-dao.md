---
title: QueryDefs.Count Property (DAO)
TOCTitle: Count Property
ms:assetid: 8caa01c5-692f-95e4-4b11-6e6c591f5872
ms:mtpsurl: https://msdn.microsoft.com/library/Ff197340(v=office.15)
ms:contentKeyID: 48546240
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: be630e54ef55fae335347dabd61440dfe457c5ed
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471084"
---
# <a name="querydefscount-property-dao"></a><span data-ttu-id="39a2a-102">QueryDefs.Count Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="39a2a-102">QueryDefs.Count Property (DAO)</span></span>


<span data-ttu-id="39a2a-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="39a2a-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="39a2a-p101">Renvoie le nombre d'objets dans la collection spécifiée. En lecture seule.</span><span class="sxs-lookup"><span data-stu-id="39a2a-p101">Returns the number of objects in the specified collection. Read-only.</span></span>

## <a name="syntax"></a><span data-ttu-id="39a2a-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="39a2a-106">Syntax</span></span>

<span data-ttu-id="39a2a-107">*expression* . Nombre</span><span class="sxs-lookup"><span data-stu-id="39a2a-107">*expression* .Count</span></span>

<span data-ttu-id="39a2a-108">*expression* Variable qui représente un objet **QueryDefs** .</span><span class="sxs-lookup"><span data-stu-id="39a2a-108">*expression* A variable that represents a **QueryDefs** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="39a2a-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="39a2a-109">Remarks</span></span>

<span data-ttu-id="39a2a-p102">Étant donné que les membres d'une collection commencent par 0, vous avez toujours intérêt à coder les boucles en commençant par le membre 0 et en finissant par la valeur de la propriété **Count** moins 1. Si vous voulez effectuer une boucle sur tous les membres d'une collection sans vérifier la propriété **Count**, vous pouvez utiliser la commande **For Each...Next**.</span><span class="sxs-lookup"><span data-stu-id="39a2a-p102">Because members of a collection begin with 0, you should always code loops starting with the 0 member and ending with the value of the **Count** property minus 1. If you want to loop through the members of a collection without checking the **Count** property, you can use a **For Each...Next** command.</span></span>

<span data-ttu-id="39a2a-p103">La valeur de la propriété **Count** n'est jamais Null. Si sa valeur est égale à 0, cela signifie qu'il n'y a pas d'objet dans la collection.</span><span class="sxs-lookup"><span data-stu-id="39a2a-p103">The **Count** property setting is never Null. If its value is 0, there are no objects in the collection.</span></span>
