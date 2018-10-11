---
title: Informations sur les erreurs relatives aux champs
TOCTitle: Field-Related Error Information
ms:assetid: 81a2c5a4-ab09-53d8-b270-e889b00a0c1a
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249559(v=office.15)
ms:contentKeyID: 48545958
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: dfa35220e463cada69c94fe56a72732591683686
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469678"
---
# <a name="field-related-error-information"></a><span data-ttu-id="77d6e-102">Informations sur les erreurs relatives aux champs</span><span class="sxs-lookup"><span data-stu-id="77d6e-102">Field-Related Error Information</span></span>


<span data-ttu-id="77d6e-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="77d6e-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="77d6e-p101">Si une erreur est directement liée à un champ  par exemple, si les données sont manquantes ou si le type du champ est incorrect  vous pouvez obtenir plus d'informations sur la cause du problème en examinant la propriété **Status** de l'objet **Field**. Cette propriété a été améliorée afin de fournir des informations spécifiques relatives au problème. Ainsi, par exemple, lorsqu'une invocation de **UpdateBatch** échoue, la cause du problème peut être déterminée en examinant la propriété **Status** des **champs** de chaque enregistrement affecté. La propriété contient une des valeurs dans la constante **FieldStatusEnum**. Le tableau suivant contient les valeurs qui présentent un intérêt lorsqu'une erreur survient.</span><span class="sxs-lookup"><span data-stu-id="77d6e-p101">If an error is directly related to a field — for example, if the data is missing or if it is the wrong type for the field — you can retrieve more information about the cause of the problem by examining the **Field** object's **Status** property. This property has been enhanced to provide specific information about the problem. So, for example, when a call to **UpdateBatch** fails, the cause of the problem can be determined by examining the **Status** property of the **Fields** in each of the effected records. The property will contain one of the values in the **FieldStatusEnum** constant. The following table includes those values that are of particular interest when an error occurs.</span></span>

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="77d6e-109">Constant</span><span class="sxs-lookup"><span data-stu-id="77d6e-109">Constant</span></span></p></th>
<th><p><span data-ttu-id="77d6e-110">Valeur</span><span class="sxs-lookup"><span data-stu-id="77d6e-110">Value</span></span></p></th>
<th><p><span data-ttu-id="77d6e-111">Description</span><span class="sxs-lookup"><span data-stu-id="77d6e-111">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="77d6e-112"><strong>adFieldCantConvertValue</strong></span><span class="sxs-lookup"><span data-stu-id="77d6e-112"><strong>adFieldCantConvertValue</strong></span></span></p></td>
<td><p><span data-ttu-id="77d6e-113">2</span><span class="sxs-lookup"><span data-stu-id="77d6e-113">2</span></span></p></td>
<td><p><span data-ttu-id="77d6e-114">Indique que le champ ne peut pas être extrait ni stocké sans perte de données.</span><span class="sxs-lookup"><span data-stu-id="77d6e-114">Indicates that the field cannot be retrieved or stored without loss of data.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="77d6e-115"><strong>adFieldDataOverflow</strong></span><span class="sxs-lookup"><span data-stu-id="77d6e-115"><strong>adFieldDataOverflow</strong></span></span></p></td>
<td><p><span data-ttu-id="77d6e-116">6</span><span class="sxs-lookup"><span data-stu-id="77d6e-116">6</span></span></p></td>
<td><p><span data-ttu-id="77d6e-117">Indique que les données renvoyées par le fournisseur ont dépassé le type de données du champ.</span><span class="sxs-lookup"><span data-stu-id="77d6e-117">Indicates that the data returned from the provider overflowed the data type of the field.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="77d6e-118"><strong>adFieldDefault</strong></span><span class="sxs-lookup"><span data-stu-id="77d6e-118"><strong>adFieldDefault</strong></span></span></p></td>
<td><p><span data-ttu-id="77d6e-119">13</span><span class="sxs-lookup"><span data-stu-id="77d6e-119">13</span></span></p></td>
<td><p><span data-ttu-id="77d6e-120">Indique que la valeur par défaut du champ a été utilisée lors de la définition des données.</span><span class="sxs-lookup"><span data-stu-id="77d6e-120">Indicates that the default value for the field was used when setting data.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="77d6e-121"><strong>adFieldIgnore</strong></span><span class="sxs-lookup"><span data-stu-id="77d6e-121"><strong>adFieldIgnore</strong></span></span></p></td>
<td><p><span data-ttu-id="77d6e-122">15</span><span class="sxs-lookup"><span data-stu-id="77d6e-122">15</span></span></p></td>
<td><p><span data-ttu-id="77d6e-p102">Indique que ce champ a été ignoré lors de la définition des valeurs de données dans la source. Aucune valeur n'est définie par le fournisseur.</span><span class="sxs-lookup"><span data-stu-id="77d6e-p102">Indicates that this field was skipped when setting data values in the source. No value was set by the provider.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="77d6e-125"><strong>adFieldIntegrityViolation</strong></span><span class="sxs-lookup"><span data-stu-id="77d6e-125"><strong>adFieldIntegrityViolation</strong></span></span></p></td>
<td><p><span data-ttu-id="77d6e-126">10</span><span class="sxs-lookup"><span data-stu-id="77d6e-126">10</span></span></p></td>
<td><p><span data-ttu-id="77d6e-127">Indique que le champ ne peut pas être modifié car il s'agit d'une entité calculée ou dérivée.</span><span class="sxs-lookup"><span data-stu-id="77d6e-127">Indicates that the field cannot be modified because it is a calculated or derived entity.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="77d6e-128"><strong>adFieldIsNull</strong></span><span class="sxs-lookup"><span data-stu-id="77d6e-128"><strong>adFieldIsNull</strong></span></span></p></td>
<td><p><span data-ttu-id="77d6e-129">3</span><span class="sxs-lookup"><span data-stu-id="77d6e-129">3</span></span></p></td>
<td><p><span data-ttu-id="77d6e-130">Indique que le fournisseur a renvoyé une valeur NULL.</span><span class="sxs-lookup"><span data-stu-id="77d6e-130">Indicates that the provider returned a null value.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="77d6e-131"><strong>adFieldOutOfSpace</strong></span><span class="sxs-lookup"><span data-stu-id="77d6e-131"><strong>adFieldOutOfSpace</strong></span></span></p></td>
<td><p><span data-ttu-id="77d6e-132">22</span><span class="sxs-lookup"><span data-stu-id="77d6e-132">22</span></span></p></td>
<td><p><span data-ttu-id="77d6e-133">Indique que le fournisseur ne parvient pas à obtenir suffisamment d'espace de stockage pour réaliser un déplacement ou une copie.</span><span class="sxs-lookup"><span data-stu-id="77d6e-133">Indicates that the provider is unable to obtain enough storage space to complete a move or copy operation.</span></span></p></td>
</tr>
</tbody>
</table>
