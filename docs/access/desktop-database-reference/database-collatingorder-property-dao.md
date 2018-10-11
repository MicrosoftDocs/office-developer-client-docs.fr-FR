---
title: Database.CollatingOrder Property (DAO)
TOCTitle: CollatingOrder Property
ms:assetid: 7f6c35bf-e5f9-8423-608e-bc072ca09141
ms:mtpsurl: https://msdn.microsoft.com/library/Ff196459(v=office.15)
ms:contentKeyID: 48545901
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 6101729af7f77e1451de240deece2bf789110fb7
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471612"
---
# <a name="databasecollatingorder-property-dao"></a><span data-ttu-id="b3f34-102">Database.CollatingOrder Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="b3f34-102">Database.CollatingOrder Property (DAO)</span></span>


<span data-ttu-id="b3f34-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="b3f34-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="b3f34-p101">Renvoie une valeur qui spécifie la séquence de l'ordre de tri du texte pour la comparaison ou le tri de chaînes de caractères (espaces de travail Microsoft Access uniquement). Valeur **Long** en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="b3f34-p101">Returns a value that specifies the sequence of the sort order in text for string comparison or sorting (Microsoft Access workspaces only). Read-only **Long**.</span></span>

## <a name="syntax"></a><span data-ttu-id="b3f34-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b3f34-106">Syntax</span></span>

<span data-ttu-id="b3f34-107">*expression* . CollatingOrder</span><span class="sxs-lookup"><span data-stu-id="b3f34-107">*expression* .CollatingOrder</span></span>

<span data-ttu-id="b3f34-108">*expression* Variable qui représente un objet de **base de données** .</span><span class="sxs-lookup"><span data-stu-id="b3f34-108">*expression* A variable that represents a **Database** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="b3f34-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="b3f34-109">Remarks</span></span>

<span data-ttu-id="b3f34-110">La valeur de retour est une valeur de type **Long** ou une constante pouvant avoir l'une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="b3f34-110">The return value is a **Long** value or constant that can be one of the following values.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="b3f34-111">Constante</span><span class="sxs-lookup"><span data-stu-id="b3f34-111">Constant</span></span></p></th>
<th><p><span data-ttu-id="b3f34-112">Ordre de tri</span><span class="sxs-lookup"><span data-stu-id="b3f34-112">Sort order</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-113"><strong>dbSortGeneral</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-113"><strong>dbSortGeneral</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-114">Général (Anglais, Français, Allemand, Portugais, Italien et Espagnol moderne)</span><span class="sxs-lookup"><span data-stu-id="b3f34-114">General (English, French, German, Portuguese, Italian, and Modern Spanish)</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-115"><strong>dbSortArabic</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-115"><strong>dbSortArabic</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-116">Arabe</span><span class="sxs-lookup"><span data-stu-id="b3f34-116">Arabic</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-117"><strong>dbSortChineseSimplified</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-117"><strong>dbSortChineseSimplified</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-118">Chinois simplifié</span><span class="sxs-lookup"><span data-stu-id="b3f34-118">Simplified Chinese</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-119"><strong>dbSortChineseTraditional</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-119"><strong>dbSortChineseTraditional</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-120">Chinois traditionnel</span><span class="sxs-lookup"><span data-stu-id="b3f34-120">Traditional Chinese</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-121"><strong>dbSortCyrillic</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-121"><strong>dbSortCyrillic</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-122">Russe</span><span class="sxs-lookup"><span data-stu-id="b3f34-122">Russian</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-123"><strong>dbSortCzech</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-123"><strong>dbSortCzech</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-124">Tchèque</span><span class="sxs-lookup"><span data-stu-id="b3f34-124">Czech</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-125"><strong>dbSortDutch</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-125"><strong>dbSortDutch</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-126">Néerlandais</span><span class="sxs-lookup"><span data-stu-id="b3f34-126">Dutch</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-127"><strong>dbSortGreek</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-127"><strong>dbSortGreek</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-128">Grec</span><span class="sxs-lookup"><span data-stu-id="b3f34-128">Greek</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-129"><strong>dbSortHebrew</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-129"><strong>dbSortHebrew</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-130">Hébreu</span><span class="sxs-lookup"><span data-stu-id="b3f34-130">Hebrew</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-131"><strong>dbSortHungarian</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-131"><strong>dbSortHungarian</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-132">Hongrois</span><span class="sxs-lookup"><span data-stu-id="b3f34-132">Hungarian</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-133"><strong>dbSortIcelandic</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-133"><strong>dbSortIcelandic</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-134">Islandais</span><span class="sxs-lookup"><span data-stu-id="b3f34-134">Icelandic</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-135"><strong>dbSortJapanese</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-135"><strong>dbSortJapanese</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-136">Japonais</span><span class="sxs-lookup"><span data-stu-id="b3f34-136">Japanese</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-137"><strong>dbSortKorean</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-137"><strong>dbSortKorean</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-138">Coréen</span><span class="sxs-lookup"><span data-stu-id="b3f34-138">Korean</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-139"><strong>dbSortNeutral</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-139"><strong>dbSortNeutral</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-140">Neutre</span><span class="sxs-lookup"><span data-stu-id="b3f34-140">Neutral</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-141"><strong>dbSortNorwDan</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-141"><strong>dbSortNorwDan</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-142">Norvégien ou Danois</span><span class="sxs-lookup"><span data-stu-id="b3f34-142">Norwegian or Danish</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-143"><strong>dbSortPDXIntl</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-143"><strong>dbSortPDXIntl</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-144">International Paradox</span><span class="sxs-lookup"><span data-stu-id="b3f34-144">Paradox International</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-145"><strong>dbSortPDXNor</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-145"><strong>dbSortPDXNor</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-146">Norvégien ou Danois Paradox</span><span class="sxs-lookup"><span data-stu-id="b3f34-146">Paradox Norwegian or Danish</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-147"><strong>dbSortPDXSwe</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-147"><strong>dbSortPDXSwe</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-148">Suédois ou Finnois Paradox</span><span class="sxs-lookup"><span data-stu-id="b3f34-148">Paradox Swedish or Finnish</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-149"><strong>dbSortPolish</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-149"><strong>dbSortPolish</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-150">Polonais</span><span class="sxs-lookup"><span data-stu-id="b3f34-150">Polish</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-151"><strong>dbSortSlovenian</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-151"><strong>dbSortSlovenian</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-152">Slovène</span><span class="sxs-lookup"><span data-stu-id="b3f34-152">Slovenian</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-153"><strong>dbSortSpanish</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-153"><strong>dbSortSpanish</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-154">Espagnol</span><span class="sxs-lookup"><span data-stu-id="b3f34-154">Spanish</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-155"><strong>dbSortSwedFin</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-155"><strong>dbSortSwedFin</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-156">Suédois ou Finnois</span><span class="sxs-lookup"><span data-stu-id="b3f34-156">Swedish or Finnish</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-157"><strong>dbSortThai</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-157"><strong>dbSortThai</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-158">Thaï</span><span class="sxs-lookup"><span data-stu-id="b3f34-158">Thai</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="b3f34-159"><strong>dbSortTurkish</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-159"><strong>dbSortTurkish</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-160">Turc</span><span class="sxs-lookup"><span data-stu-id="b3f34-160">Turkish</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="b3f34-161"><strong>dbSortUndefined</strong></span><span class="sxs-lookup"><span data-stu-id="b3f34-161"><strong>dbSortUndefined</strong></span></span></p></td>
<td><p><span data-ttu-id="b3f34-162">Non défini ou inconnu</span><span class="sxs-lookup"><span data-stu-id="b3f34-162">Undefined or unknown</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="b3f34-163">Le paramètre de propriété **CollatingOrder** correspond à l’argument paramètres régionaux de la méthode **CreateDatabase** lors de la création de la base de données ou de la méthode **CompactDatabase** lors du dernier compactage de la base de données.</span><span class="sxs-lookup"><span data-stu-id="b3f34-163">The **CollatingOrder** property setting corresponds to the locale argument of the **CreateDatabase** method when the database was created or the **CompactDatabase** method when the database was most recently compacted.</span></span>

<span data-ttu-id="b3f34-164">Vérifiez le paramètre de propriété **CollatingOrder** d’un objet de **base de données** ou de **champ** pour déterminer la méthode de comparaison de chaîne pour la base de données ou d’un champ.</span><span class="sxs-lookup"><span data-stu-id="b3f34-164">Check the **CollatingOrder** property setting of a **Database** or **Field** object to determine the string comparison method for the database or field.</span></span> <span data-ttu-id="b3f34-165">Vous pouvez définir la propriété **CollatingOrder** d’un nouvel objet **Field** non ajouté si vous souhaitez que le paramètre de l’objet **Field** diffère de celle de l’objet **Database** qui le contient.</span><span class="sxs-lookup"><span data-stu-id="b3f34-165">You can set the **CollatingOrder** property of a new, unappended **Field** object if you want the setting of the **Field** object to differ from that of the **Database** object that contains it.</span></span>
