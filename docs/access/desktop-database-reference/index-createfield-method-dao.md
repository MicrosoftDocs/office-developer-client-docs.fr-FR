---
title: Index.CreateField Method (DAO)
TOCTitle: CreateField Method
ms:assetid: fc82b785-8768-b144-a2a4-c1f1798865a6
ms:mtpsurl: https://msdn.microsoft.com/library/Ff837208(v=office.15)
ms:contentKeyID: 48548892
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: a3723beee3907cc209951c3494a2bde2a4c17bfa
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470160"
---
# <a name="indexcreatefield-method-dao"></a><span data-ttu-id="6e0ff-102">Index.CreateField Method (DAO)</span><span class="sxs-lookup"><span data-stu-id="6e0ff-102">Index.CreateField Method (DAO)</span></span>


<span data-ttu-id="6e0ff-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="6e0ff-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="6e0ff-104">Crée un objet **[Field](field-object-dao.md)** (espaces de travail Microsoft Access uniquement).</span><span class="sxs-lookup"><span data-stu-id="6e0ff-104">Creates a new **[Field](field-object-dao.md)** object (Microsoft Access workspaces only).</span></span>

## <a name="syntax"></a><span data-ttu-id="6e0ff-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="6e0ff-105">Syntax</span></span>

<span data-ttu-id="6e0ff-106">*expression* . CreateField (***nom***, ***Type***, ***taille***)</span><span class="sxs-lookup"><span data-stu-id="6e0ff-106">*expression* .CreateField(***Name***, ***Type***, ***Size***)</span></span>

<span data-ttu-id="6e0ff-107">*expression* Variable qui représente un objet **Index** .</span><span class="sxs-lookup"><span data-stu-id="6e0ff-107">*expression* A variable that represents an **Index** object.</span></span>

### <a name="parameters"></a><span data-ttu-id="6e0ff-108">Paramètres</span><span class="sxs-lookup"><span data-stu-id="6e0ff-108">Parameters</span></span>

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="6e0ff-109">Name</span><span class="sxs-lookup"><span data-stu-id="6e0ff-109">Name</span></span></p></th>
<th><p><span data-ttu-id="6e0ff-110">Obligatoire/Facultatif</span><span class="sxs-lookup"><span data-stu-id="6e0ff-110">Required/Optional</span></span></p></th>
<th><p><span data-ttu-id="6e0ff-111">Type de données</span><span class="sxs-lookup"><span data-stu-id="6e0ff-111">Data Type</span></span></p></th>
<th><p><span data-ttu-id="6e0ff-112">Description</span><span class="sxs-lookup"><span data-stu-id="6e0ff-112">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="6e0ff-113">Name</span><span class="sxs-lookup"><span data-stu-id="6e0ff-113">Name</span></span></p></td>
<td><p><span data-ttu-id="6e0ff-114">Facultatif</span><span class="sxs-lookup"><span data-stu-id="6e0ff-114">Optional</span></span></p></td>
<td><p><span data-ttu-id="6e0ff-115"><strong>Variant</strong></span><span class="sxs-lookup"><span data-stu-id="6e0ff-115"><strong>Variant</strong></span></span></p></td>
<td><p><span data-ttu-id="6e0ff-p101">Chaîne qui identifie de manière unique le nouvel objet <strong>Field</strong>. Reportez-vous à la propriété <strong><a href="connection-name-property-dao.md">Name</a></strong> pour plus d'informations sur les noms valides pour l'objet <strong>Field</strong>.  </span><span class="sxs-lookup"><span data-stu-id="6e0ff-p101">A String that uniquely names the new <strong>Field</strong> object. See the <strong><a href="connection-name-property-dao.md">Name</a></strong> property for details on valid <strong>Field</strong> names.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="6e0ff-118">Type</span><span class="sxs-lookup"><span data-stu-id="6e0ff-118">Type</span></span></p></td>
<td><p><span data-ttu-id="6e0ff-119">Facultatif</span><span class="sxs-lookup"><span data-stu-id="6e0ff-119">Optional</span></span></p></td>
<td><p><span data-ttu-id="6e0ff-120"><strong>Variante</strong></span><span class="sxs-lookup"><span data-stu-id="6e0ff-120"><strong>Variant</strong></span></span></p></td>
<td><p><span data-ttu-id="6e0ff-121">Argument non pris en charge pour cet objet.</span><span class="sxs-lookup"><span data-stu-id="6e0ff-121">Argument not supported for this object.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="6e0ff-122">Size</span><span class="sxs-lookup"><span data-stu-id="6e0ff-122">Size</span></span></p></td>
<td><p><span data-ttu-id="6e0ff-123">Facultatif</span><span class="sxs-lookup"><span data-stu-id="6e0ff-123">Optional</span></span></p></td>
<td><p><span data-ttu-id="6e0ff-124"><strong>Variante</strong></span><span class="sxs-lookup"><span data-stu-id="6e0ff-124"><strong>Variant</strong></span></span></p></td>
<td><p><span data-ttu-id="6e0ff-125">Argument non pris en charge pour cet objet.</span><span class="sxs-lookup"><span data-stu-id="6e0ff-125">Argument not supported for this object.</span></span></p></td>
</tr>
</tbody>
</table>


### <a name="return-value"></a><span data-ttu-id="6e0ff-126">Valeur renvoyée</span><span class="sxs-lookup"><span data-stu-id="6e0ff-126">Return Value</span></span>

<span data-ttu-id="6e0ff-127">Champ</span><span class="sxs-lookup"><span data-stu-id="6e0ff-127">Field</span></span>

## <a name="remarks"></a><span data-ttu-id="6e0ff-128">Remarques</span><span class="sxs-lookup"><span data-stu-id="6e0ff-128">Remarks</span></span>

<span data-ttu-id="6e0ff-p102">Vous pouvez utiliser la méthode **CreateField** pour créer un champ, spécifier le nom, le type de données et la taille du champ. Si vous omettez une ou plusieurs des parties facultatives lorsque vous utilisez la méthode **CreateField**, vous pouvez utiliser une instruction d'affectation appropriée pour définir ou réinitialiser la propriété correspondante avant d'ajouter le nouvel objet à la collection. Une fois que vous avez ajouté le nouvel objet, vous pouvez modifier une partie de ses paramètres de propriété, mais pas tous. Pour plus d'informations, reportez-vous aux rubriques concernant cette propriété.</span><span class="sxs-lookup"><span data-stu-id="6e0ff-p102">You can use the **CreateField** method to create a new field, as well as specify the name, data type, and size of the field. If you omit one or more of the optional parts when you use **CreateField**, you can use an appropriate assignment statement to set or reset the corresponding property before you append the new object to a collection. After you append the new object, you can alter some but not all of its property settings. See the individual property topics for more details.</span></span>

<span data-ttu-id="6e0ff-133">Les arguments de type et la taille s’appliquent uniquement aux objets **Field** dans un objet **TableDef** .</span><span class="sxs-lookup"><span data-stu-id="6e0ff-133">The type and size arguments apply only to **Field** objects in a **TableDef** object.</span></span> <span data-ttu-id="6e0ff-134">Ces arguments sont ignorés lorsqu'un objet **Field** est associé à un objet **Index** ou **Relation**.</span><span class="sxs-lookup"><span data-stu-id="6e0ff-134">These arguments are ignored when a **Field** object is associated with an **Index** or **Relation** object.</span></span>

<span data-ttu-id="6e0ff-135">Si le nom fait référence à un objet qui est déjà membre de la collection, une erreur d’exécution se produit lorsque vous utilisez la méthode **[Append](fields-append-method-dao.md)** .</span><span class="sxs-lookup"><span data-stu-id="6e0ff-135">If name refers to an object that is already a member of the collection, a run-time error occurs when you use the **[Append](fields-append-method-dao.md)** method.</span></span>

<span data-ttu-id="6e0ff-p104">Pour supprimer un objet **Field** d'une collection **Fields**, utilisez la méthode **[Delete](fields-delete-method-dao.md)** dans la collection. Vous ne pouvez pas supprimer un objet **Field** dans la collection **Fields** d'un objet **TableDef** une fois que vous avez créé un index qui renvoie à ce champ.</span><span class="sxs-lookup"><span data-stu-id="6e0ff-p104">To remove a **Field** object from a **Fields** collection, use the **[Delete](fields-delete-method-dao.md)** method on the collection. You can't delete a **Field** object from a **TableDef** object's **Fields** collection after you create an index that references the field.</span></span>
