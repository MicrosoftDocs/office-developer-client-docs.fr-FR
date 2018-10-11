---
title: Recordset.MoveLast Method (DAO)
TOCTitle: MoveLast Method
ms:assetid: fc0f7a33-1f55-9f5b-b00d-1b81f49b1c3e
ms:mtpsurl: https://msdn.microsoft.com/library/Ff837192(v=office.15)
ms:contentKeyID: 48548881
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: c154a576a4ac1aedf555ba8164bc243fe49a3841
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471287"
---
# <a name="recordsetmovelast-method-dao"></a><span data-ttu-id="acd71-102">Recordset.MoveLast Method (DAO)</span><span class="sxs-lookup"><span data-stu-id="acd71-102">Recordset.MoveLast Method (DAO)</span></span>


<span data-ttu-id="acd71-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="acd71-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="acd71-104">Atteint le dernier enregistrement d'un objet **Recordset** spécifié et en fait l'enregistrement actif.</span><span class="sxs-lookup"><span data-stu-id="acd71-104">Moves to the last record in a specified **Recordset** object and make that record the current record.</span></span>

## <a name="syntax"></a><span data-ttu-id="acd71-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="acd71-105">Syntax</span></span>

<span data-ttu-id="acd71-106">*expression* . MoveLast (***Options***)</span><span class="sxs-lookup"><span data-stu-id="acd71-106">*expression* .MoveLast(***Options***)</span></span>

<span data-ttu-id="acd71-107">*expression* Variable qui représente un objet **Recordset** .</span><span class="sxs-lookup"><span data-stu-id="acd71-107">*expression* A variable that represents a **Recordset** object.</span></span>

### <a name="parameters"></a><span data-ttu-id="acd71-108">Paramètres</span><span class="sxs-lookup"><span data-stu-id="acd71-108">Parameters</span></span>

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="acd71-109">Name</span><span class="sxs-lookup"><span data-stu-id="acd71-109">Name</span></span></p></th>
<th><p><span data-ttu-id="acd71-110">Obligatoire/Facultatif</span><span class="sxs-lookup"><span data-stu-id="acd71-110">Required/Optional</span></span></p></th>
<th><p><span data-ttu-id="acd71-111">Type de données</span><span class="sxs-lookup"><span data-stu-id="acd71-111">Data Type</span></span></p></th>
<th><p><span data-ttu-id="acd71-112">Description</span><span class="sxs-lookup"><span data-stu-id="acd71-112">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="acd71-113">Options</span><span class="sxs-lookup"><span data-stu-id="acd71-113">Options</span></span></p></td>
<td><p><span data-ttu-id="acd71-114">Facultatif</span><span class="sxs-lookup"><span data-stu-id="acd71-114">Optional</span></span></p></td>
<td><p><span data-ttu-id="acd71-115"><strong>Entier long</strong></span><span class="sxs-lookup"><span data-stu-id="acd71-115"><strong>Long</strong></span></span></p></td>
<td><p><span data-ttu-id="acd71-116">Associez ce paramètre à la constante <strong>dbRunAsync</strong> pour obtenir une exécution asynchrone de l'appel de <strong>MoveLast</strong></span><span class="sxs-lookup"><span data-stu-id="acd71-116">Set to <strong>dbRunAsync</strong> to rune the call to <strong>MoveLast</strong> asynchronously.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="acd71-117">Remarques</span><span class="sxs-lookup"><span data-stu-id="acd71-117">Remarks</span></span>

<span data-ttu-id="acd71-118">Utilisez les méthodes **Move** pour vous déplacer entre les enregistrements sans appliquer de condition.</span><span class="sxs-lookup"><span data-stu-id="acd71-118">Use the **Move** methods to move from record to record without applying a condition.</span></span>

<span data-ttu-id="acd71-p101">Si vous modifiez l'enregistrement actif, veillez à utiliser la méthode **Update** pour enregistrer les modifications avant de passer à un autre enregistrement. Si vous accédez à un autre enregistrement alors que vous effectuez une mise à jour, les modifications seront perdues sans notification.</span><span class="sxs-lookup"><span data-stu-id="acd71-p101">If you edit the current record, be sure you use the **Update** method to save the changes before you move to another record. If you move to another record without updating, your changes are lost without warning.</span></span>

<span data-ttu-id="acd71-p102">Lorsque vous ouvrez un objet **Recordset**, le premier enregistrement est actif et la propriété **BOF** a la valeur **False**. Si l'objet **Recordset** ne contient pas d'enregistrements, la propriété **BOF** a la valeur **True** et aucun enregistrement n'est actif.</span><span class="sxs-lookup"><span data-stu-id="acd71-p102">When you open a **Recordset**, the first record is current and the **BOF** property is **False**. If the **Recordset** contains no records, the **BOF** property is **True**, and there is no current record.</span></span>

<span data-ttu-id="acd71-123">Si le premier ou le dernier enregistrement est déjà actif lorsque vous utilisez la méthode **MoveFirst** ou **MoveLast**, l'enregistrement actif ne change pas.</span><span class="sxs-lookup"><span data-stu-id="acd71-123">If the first or last record is already current when you use **MoveFirst** or **MoveLast**, the current record doesn't change.</span></span>

<span data-ttu-id="acd71-124">Si le jeu d’enregistrements fait référence à un **objet Recordset** de type table (espaces de travail Microsoft Access uniquement), le déplacement suit l’index actuel.</span><span class="sxs-lookup"><span data-stu-id="acd71-124">If recordset refers to a table-type **Recordset** (Microsoft Access workspaces only), movement follows the current index.</span></span> <span data-ttu-id="acd71-125">Vous pouvez définir l'index actif à l'aide de la propriété **Index**.</span><span class="sxs-lookup"><span data-stu-id="acd71-125">You can set the current index by using the **Index** property.</span></span> <span data-ttu-id="acd71-126">Si vous ne définissez pas l'index actuel, l'ordre des enregistrements renvoyés est indéfini.</span><span class="sxs-lookup"><span data-stu-id="acd71-126">If you don't set the current index, the order of returned records is undefined.</span></span>


> [!NOTE]
> <P><span data-ttu-id="acd71-p104">[!REMARQUE] Vous pouvez utiliser la méthode <STRONG>MoveLast</STRONG> pour remplir entièrement un objet <STRONG>Recordset</STRONG> de type feuille de réponse dynamique ou instantané en vue de fournir le nombre actuel d'enregistrements présents dans l'objet <STRONG>Recordset</STRONG>. Toutefois, si vous utilisez <STRONG>MoveLast</STRONG> de cette manière, vous risquez de ralentir l'exécution de votre application. Il n'est conseillé d'utiliser la méthode <STRONG>MoveLast</STRONG> qu'en cas de nécessité absolue d'obtenir un décompte précis des enregistrements présents dans un objet <STRONG>Recordset</STRONG> récemment ouvert. Si vous utilisez la constante <STRONG>dbRunAsync</STRONG> avec <STRONG>MoveLast</STRONG>, l'appel de la méthode est asynchrone. Vous pouvez utiliser la propriété <STRONG>StillExecuting</STRONG> pour déterminer à quel moment l'enregistrement <STRONG>Recordset</STRONG> est rempli entièrement, de même que vous pouvez utiliser la méthode <STRONG>Cancel</STRONG> pour terminer l'exécution de l'appel asynchrone de la méthode <STRONG>MoveLast</STRONG>.</span><span class="sxs-lookup"><span data-stu-id="acd71-p104">You can use the <STRONG>MoveLast</STRONG> method to fully populate a dynaset- or snapshot-type <STRONG>Recordset</STRONG> to provide the current number of records in the <STRONG>Recordset</STRONG>. However, if you use <STRONG>MoveLast</STRONG> in this way, you can slow down your application's performance. You should only use <STRONG>MoveLast</STRONG> to get a record count if it is absolutely necessary to obtain an accurate record count on a newly opened <STRONG>Recordset</STRONG>. If you use the <STRONG>dbRunAsync</STRONG> constant with <STRONG>MoveLast</STRONG>, the method call is asynchronous. You can use the <STRONG>StillExecuting</STRONG> property to determine when the <STRONG>Recordset</STRONG> is fully populated, and you can use the <STRONG>Cancel</STRONG> method to terminate execution of the asynchronous <STRONG>MoveLast</STRONG> method call.</span></span></P>



<span data-ttu-id="acd71-132">Vous ne pouvez pas utiliser les méthodes **MoveFirst**, **MoveLast**et **MovePrevious** sur un objet **Recordset** de type avant uniquement.</span><span class="sxs-lookup"><span data-stu-id="acd71-132">You can't use the **MoveFirst**, **MoveLast**, and **MovePrevious** methods on a forward–only–type **Recordset** object.</span></span>

<span data-ttu-id="acd71-133">Pour déplacer la position de l'enregistrement actif dans un objet **Recordset** d'un objet d'un nombre d'enregistrements spécifique vers l'avant ou l'arrière, utilisez la méthode **Move**.</span><span class="sxs-lookup"><span data-stu-id="acd71-133">To move the position of the current record in a **Recordset** object a specific number of records forward or backward, use the **Move** method.</span></span>
