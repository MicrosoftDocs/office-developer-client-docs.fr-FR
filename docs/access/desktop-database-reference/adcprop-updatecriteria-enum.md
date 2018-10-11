---
title: ADCPROP_UPDATECRITERIA_ENUM
TOCTitle: ADCPROP_UPDATECRITERIA_ENUM
ms:assetid: 70da63fa-fa75-9bb4-683d-0fcb4c4a2934
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249450(v=office.15)
ms:contentKeyID: 48545571
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: b018402f6a2d48819a3f97c69e2a8db77fa9e84c
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472549"
---
# <a name="adcpropupdatecriteriaenum"></a><span data-ttu-id="f4570-102">ADCPROP\_UPDATECRITERIA\_ENUM</span><span class="sxs-lookup"><span data-stu-id="f4570-102">ADCPROP\_UPDATECRITERIA\_ENUM</span></span>

<span data-ttu-id="f4570-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="f4570-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="f4570-104">Indique les champs qui peuvent être utilisés pour détecter les conflits pendant une mise à jour optimiste d'une ligne de la source de données avec un objet [Recordset](recordset-object-ado.md).</span><span class="sxs-lookup"><span data-stu-id="f4570-104">Specifies which fields can be used to detect conflicts during an optimistic update of a row of the data source with a [Recordset](recordset-object-ado.md) object.</span></span>

<span data-ttu-id="f4570-105">Utilisez ces constantes avec la propriété dynamique "**Update Criteria**" du **Recordset**, qui est référencée dans l’[index des propriétés dynamiques ADO](ado-dynamic-property-index.md) et expliquée dans la documentation [Microsoft Cursor Service for OLE DB](microsoft-cursor-service-for-ole-db-ado-service-component.md).</span><span class="sxs-lookup"><span data-stu-id="f4570-105">Use these constants with the **Recordset** "**Update Criteria**" dynamic property, which is referenced in the [ADO Dynamic Property Index](ado-dynamic-property-index.md) and documented in the [Microsoft Cursor Service for OLE DB](microsoft-cursor-service-for-ole-db-ado-service-component.md) documentation.</span></span>

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="f4570-106">Constant</span><span class="sxs-lookup"><span data-stu-id="f4570-106">Constant</span></span></p></th>
<th><p><span data-ttu-id="f4570-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="f4570-107">Value</span></span></p></th>
<th><p><span data-ttu-id="f4570-108">Description</span><span class="sxs-lookup"><span data-stu-id="f4570-108">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="f4570-109"><strong>adCriteriaAllCols</strong></span><span class="sxs-lookup"><span data-stu-id="f4570-109"><strong>adCriteriaAllCols</strong></span></span></p></td>
<td><p><span data-ttu-id="f4570-110">1</span><span class="sxs-lookup"><span data-stu-id="f4570-110">1</span></span></p></td>
<td><p><span data-ttu-id="f4570-111">Détecte des conflits en cas de modification d'une ligne de la source de données.</span><span class="sxs-lookup"><span data-stu-id="f4570-111">Detects conflicts if any column of the data source row has been changed.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="f4570-112"><strong>adCriteriaKey</strong></span><span class="sxs-lookup"><span data-stu-id="f4570-112"><strong>adCriteriaKey</strong></span></span></p></td>
<td><p><span data-ttu-id="f4570-113">0</span><span class="sxs-lookup"><span data-stu-id="f4570-113">0</span></span></p></td>
<td><p><span data-ttu-id="f4570-114">Détecte des conflits en cas de modification de la colonne de clés de la ligne de source de données, indiquant que la ligne a été supprimée.</span><span class="sxs-lookup"><span data-stu-id="f4570-114">Detects conflicts if the key column of the data source row has been changed, which means that the row has been deleted.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="f4570-115"><strong>adCriteriaTimeStamp</strong></span><span class="sxs-lookup"><span data-stu-id="f4570-115"><strong>adCriteriaTimeStamp</strong></span></span></p></td>
<td><p><span data-ttu-id="f4570-116">3</span><span class="sxs-lookup"><span data-stu-id="f4570-116">3</span></span></p></td>
<td><p><span data-ttu-id="f4570-117">Détecte des conflits si l'horodatage de la ligne de source de données a été modifié, signifiant qu'il y a eu un accès après obtention du <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="f4570-117">Detects conflicts if the timestamp of the data source row has been changed, which means the row has been accessed after the <strong>Recordset</strong> was obtained.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="f4570-118"><strong>adCriteriaUpdCols</strong></span><span class="sxs-lookup"><span data-stu-id="f4570-118"><strong>adCriteriaUpdCols</strong></span></span></p></td>
<td><p><span data-ttu-id="f4570-119">2</span><span class="sxs-lookup"><span data-stu-id="f4570-119">2</span></span></p></td>
<td><p><span data-ttu-id="f4570-120">Détecte des conflits si une des colonnes de la ligne de source de données correspondant aux champs mis à jour du <strong>Recordset</strong> a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="f4570-120">Detects conflicts if any of the columns of the data source row that correspond to updated fields of the <strong>Recordset</strong> have been changed.</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="f4570-121">**Équivalent ADO/WFC**</span><span class="sxs-lookup"><span data-stu-id="f4570-121">**ADO/WFC Equivalent**</span></span>

<span data-ttu-id="f4570-122">Module : **com.ms.wfc.data**</span><span class="sxs-lookup"><span data-stu-id="f4570-122">Package: **com.ms.wfc.data**</span></span>

<table>
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="f4570-123">Constante</span><span class="sxs-lookup"><span data-stu-id="f4570-123">Constant</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="f4570-124">AdoEnums.AdcPropUpdateCriteria.ALLCOLS</span><span class="sxs-lookup"><span data-stu-id="f4570-124">AdoEnums.AdcPropUpdateCriteria.ALLCOLS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="f4570-125">AdoEnums.AdcPropUpdateCriteria.KEY</span><span class="sxs-lookup"><span data-stu-id="f4570-125">AdoEnums.AdcPropUpdateCriteria.KEY</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="f4570-126">AdoEnums.AdcPropUpdateCriteria.TIMESTAMP</span><span class="sxs-lookup"><span data-stu-id="f4570-126">AdoEnums.AdcPropUpdateCriteria.TIMESTAMP</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="f4570-127">AdoEnums.AdcPropUpdateCriteria.UPDCOLS</span><span class="sxs-lookup"><span data-stu-id="f4570-127">AdoEnums.AdcPropUpdateCriteria.UPDCOLS</span></span></p></td>
</tr>
</tbody>
</table>
