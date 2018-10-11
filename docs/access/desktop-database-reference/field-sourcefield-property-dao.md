---
title: Field.SourceField Property (DAO)
TOCTitle: SourceField Property
ms:assetid: e5750d6c-4078-7bbb-9356-f9207c4e8028
ms:mtpsurl: https://msdn.microsoft.com/library/Ff835953(v=office.15)
ms:contentKeyID: 48548360
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 729a858ae39c6403d8ccefcacaa7bade754f0484
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470501"
---
# <a name="fieldsourcefield-property-dao"></a><span data-ttu-id="e7d38-102">Field.SourceField Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="e7d38-102">Field.SourceField Property (DAO)</span></span>


<span data-ttu-id="e7d38-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="e7d38-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="e7d38-p101">Renvoie une valeur spécifiant le nom du champ qui représente la source d'origine des données d'un objet **Field**. Type **String** en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="e7d38-p101">Returns a value that indicates the name of the field that is the original source of the data for a **Field** object. Read-only **String**.</span></span>

## <a name="syntax"></a><span data-ttu-id="e7d38-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="e7d38-106">Syntax</span></span>

<span data-ttu-id="e7d38-107">*expression* . SourceField</span><span class="sxs-lookup"><span data-stu-id="e7d38-107">*expression* .SourceField</span></span>

<span data-ttu-id="e7d38-108">*expression* Variable qui représente un objet **Field** .</span><span class="sxs-lookup"><span data-stu-id="e7d38-108">*expression* A variable that represents a **Field** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="e7d38-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="e7d38-109">Remarks</span></span>

<span data-ttu-id="e7d38-110">Pour un objet **Field**, l'utilisation des propriétés **SourceField** et **SourceTable** dépend de l'objet contenant la collection **Fields** à laquelle l'objet **Field** est ajouté, comme illustré dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="e7d38-110">For a **Field** object, use of the **SourceField** and **SourceTable** properties depends on the object that contains the **Fields** collection that the **Field** object is appended to, as shown in the following table.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="e7d38-111">Objet ajouté à</span><span class="sxs-lookup"><span data-stu-id="e7d38-111">Object appended to</span></span></p></th>
<th><p><span data-ttu-id="e7d38-112">Utilisation</span><span class="sxs-lookup"><span data-stu-id="e7d38-112">Usage</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="e7d38-113"><strong>Index</strong></span><span class="sxs-lookup"><span data-stu-id="e7d38-113"><strong>Index</strong></span></span></p></td>
<td><p><span data-ttu-id="e7d38-114">Non pris en charge</span><span class="sxs-lookup"><span data-stu-id="e7d38-114">Not supported</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="e7d38-115"><strong>Objet QueryDef</strong></span><span class="sxs-lookup"><span data-stu-id="e7d38-115"><strong>QueryDef</strong></span></span></p></td>
<td><p><span data-ttu-id="e7d38-116">Lecture seule</span><span class="sxs-lookup"><span data-stu-id="e7d38-116">Read-only</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="e7d38-117"><strong>Recordset</strong></span><span class="sxs-lookup"><span data-stu-id="e7d38-117"><strong>Recordset</strong></span></span></p></td>
<td><p><span data-ttu-id="e7d38-118">Lecture seule</span><span class="sxs-lookup"><span data-stu-id="e7d38-118">Read-only</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="e7d38-119"><strong>Relation</strong></span><span class="sxs-lookup"><span data-stu-id="e7d38-119"><strong>Relation</strong></span></span></p></td>
<td><p><span data-ttu-id="e7d38-120">Non pris en charge</span><span class="sxs-lookup"><span data-stu-id="e7d38-120">Not supported</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="e7d38-121"><strong>TableDef</strong></span><span class="sxs-lookup"><span data-stu-id="e7d38-121"><strong>TableDef</strong></span></span></p></td>
<td><p><span data-ttu-id="e7d38-122">Lecture seule</span><span class="sxs-lookup"><span data-stu-id="e7d38-122">Read-only</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="e7d38-p102">Ces propriétés indiquent les noms des tables et des champs d'origine associés à un objet **Field**. Vous pouvez, par exemple, utiliser ces propriétés pour déterminer la source d'origine des données dans un champ de requête dont le nom n'a aucun rapport avec le nom du champ dans la table sous-jacente.</span><span class="sxs-lookup"><span data-stu-id="e7d38-p102">These properties indicate the original field and table names associated with a **Field** object. For example, you could use these properties to determine the original source of the data in a query field whose name is unrelated to the name of the field in the underlying table.</span></span>
