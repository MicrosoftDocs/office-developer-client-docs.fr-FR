---
title: Database.Version Property (DAO)
TOCTitle: Version Property
ms:assetid: 40faaa0c-e764-e968-f606-7e06ded80c3f
ms:mtpsurl: https://msdn.microsoft.com/library/Ff192887(v=office.15)
ms:contentKeyID: 48544440
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 4bcb360caa71131f4892409805c7afe3ceb2dcce
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472426"
---
# <a name="databaseversion-property-dao"></a><span data-ttu-id="de4c8-102">Database.Version Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="de4c8-102">Database.Version Property (DAO)</span></span>


<span data-ttu-id="de4c8-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="de4c8-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="de4c8-p101">Dans un espace de travail Microsoft Access, renvoie la version du moteur de base de données Microsoft Jet ou Microsoft Access qui a créé la base de données. Valeur **String** en lecture seule</span><span class="sxs-lookup"><span data-stu-id="de4c8-p101">In a Microsoft Access workspace, returns the vesion of the Microsoft Jet or Microsoft Access database engine that created the database. Read-only **String**.</span></span>

## <a name="syntax"></a><span data-ttu-id="de4c8-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="de4c8-106">Syntax</span></span>

<span data-ttu-id="de4c8-107">*expression* . Version</span><span class="sxs-lookup"><span data-stu-id="de4c8-107">*expression* .Version</span></span>

<span data-ttu-id="de4c8-108">*expression* Variable qui représente un objet de **base de données** .</span><span class="sxs-lookup"><span data-stu-id="de4c8-108">*expression* A variable that represents a **Database** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="de4c8-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="de4c8-109">Remarks</span></span>

<span data-ttu-id="de4c8-110">La valeur de retour est une chaîne (type de données String) qui est évaluée à un numéro de version et mise en forme comme suit.</span><span class="sxs-lookup"><span data-stu-id="de4c8-110">The return value is a String that evaluates to a version number, formatted as follows.</span></span>

  - <span data-ttu-id="de4c8-p102">Espace de travail Microsoft Access : numéro de version sous la forme « *version.inter* ». Par exemple, « 3.0 ». Le numéro de version du produit est constitué du numéro de version (3), d’un point et du numéro de version intermédiaire (0).</span><span class="sxs-lookup"><span data-stu-id="de4c8-p102">Microsoft Access workspace represents the version number in the form "*major.minor*". For example, "3.0". The product version number consists of the version number (3), a period, and the release number (0).</span></span>

<span data-ttu-id="de4c8-114">Le tableau suivant montre la version du moteur de base de données incluse dans différentes versions de produits Microsoft.</span><span class="sxs-lookup"><span data-stu-id="de4c8-114">The following table shows which version of the database engine was included with various versions of Microsoft products.</span></span>

<table style="width:100%;">
<colgroup>
<col style="width: 16%" />
<col style="width: 16%" />
<col style="width: 16%" />
<col style="width: 16%" />
<col style="width: 16%" />
<col style="width: 16%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="de4c8-115">Moteur de base de données</span><span class="sxs-lookup"><span data-stu-id="de4c8-115">Database Engine</span></span></p></th>
<th><p><span data-ttu-id="de4c8-116">Version (année de sortie)</span><span class="sxs-lookup"><span data-stu-id="de4c8-116">Version (year released)</span></span></p></th>
<th><p><span data-ttu-id="de4c8-117">Microsoft Access</span><span class="sxs-lookup"><span data-stu-id="de4c8-117">Microsoft Access</span></span></p></th>
<th><p><span data-ttu-id="de4c8-118">Microsoft Visual Basic</span><span class="sxs-lookup"><span data-stu-id="de4c8-118">Microsoft Visual Basic</span></span></p></th>
<th><p><span data-ttu-id="de4c8-119">Microsoft Excel</span><span class="sxs-lookup"><span data-stu-id="de4c8-119">Microsoft Excel</span></span></p></th>
<th><p><span data-ttu-id="de4c8-120">Microsoft Visual C++</span><span class="sxs-lookup"><span data-stu-id="de4c8-120">Microsoft Visual C++</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="de4c8-121">Microsoft Jet</span><span class="sxs-lookup"><span data-stu-id="de4c8-121">Microsoft Jet</span></span></p></td>
<td><p><span data-ttu-id="de4c8-122">1.0 (1992)</span><span class="sxs-lookup"><span data-stu-id="de4c8-122">1.0 (1992)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-123">1.0</span><span class="sxs-lookup"><span data-stu-id="de4c8-123">1.0</span></span></p></td>
<td><p><span data-ttu-id="de4c8-124">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-124">N/A</span></span></p></td>
<td><p><span data-ttu-id="de4c8-125">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-125">N/A</span></span></p></td>
<td><p><span data-ttu-id="de4c8-126">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-126">N/A</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="de4c8-127">Microsoft Jet</span><span class="sxs-lookup"><span data-stu-id="de4c8-127">Microsoft Jet</span></span></p></td>
<td><p><span data-ttu-id="de4c8-128">1.1 (1993)</span><span class="sxs-lookup"><span data-stu-id="de4c8-128">1.1 (1993)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-129">1.1</span><span class="sxs-lookup"><span data-stu-id="de4c8-129">1.1</span></span></p></td>
<td><p><span data-ttu-id="de4c8-130">3.0</span><span class="sxs-lookup"><span data-stu-id="de4c8-130">3.0</span></span></p></td>
<td><p><span data-ttu-id="de4c8-131">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-131">N/A</span></span></p></td>
<td><p><span data-ttu-id="de4c8-132">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-132">N/A</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="de4c8-133">Microsoft Jet</span><span class="sxs-lookup"><span data-stu-id="de4c8-133">Microsoft Jet</span></span></p></td>
<td><p><span data-ttu-id="de4c8-134">2.0 (1994)</span><span class="sxs-lookup"><span data-stu-id="de4c8-134">2.0 (1994)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-135">2.0</span><span class="sxs-lookup"><span data-stu-id="de4c8-135">2.0</span></span></p></td>
<td><p><span data-ttu-id="de4c8-136">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-136">N/A</span></span></p></td>
<td><p><span data-ttu-id="de4c8-137">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-137">N/A</span></span></p></td>
<td><p><span data-ttu-id="de4c8-138">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-138">N/A</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="de4c8-139">Microsoft Jet</span><span class="sxs-lookup"><span data-stu-id="de4c8-139">Microsoft Jet</span></span></p></td>
<td><p><span data-ttu-id="de4c8-140">2.5 (1995)</span><span class="sxs-lookup"><span data-stu-id="de4c8-140">2.5 (1995)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-141">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-141">N/A</span></span></p></td>
<td><p><span data-ttu-id="de4c8-142">4.0 (16 bits)</span><span class="sxs-lookup"><span data-stu-id="de4c8-142">4.0 (16-bit)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-143">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-143">N/A</span></span></p></td>
<td><p><span data-ttu-id="de4c8-144">S/O</span><span class="sxs-lookup"><span data-stu-id="de4c8-144">N/A</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="de4c8-145">Microsoft Jet</span><span class="sxs-lookup"><span data-stu-id="de4c8-145">Microsoft Jet</span></span></p></td>
<td><p><span data-ttu-id="de4c8-146">3.0 (1995)</span><span class="sxs-lookup"><span data-stu-id="de4c8-146">3.0 (1995)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-147">' 95 (7.0)</span><span class="sxs-lookup"><span data-stu-id="de4c8-147">‘95 (7.0)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-148">4.0 (32 bits)</span><span class="sxs-lookup"><span data-stu-id="de4c8-148">4.0 (32-bit)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-149">' 95 (7.0)</span><span class="sxs-lookup"><span data-stu-id="de4c8-149">‘95 (7.0)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-150">4.x</span><span class="sxs-lookup"><span data-stu-id="de4c8-150">4.x</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="de4c8-151">Microsoft Jet</span><span class="sxs-lookup"><span data-stu-id="de4c8-151">Microsoft Jet</span></span></p></td>
<td><p><span data-ttu-id="de4c8-152">3.5 (1996)</span><span class="sxs-lookup"><span data-stu-id="de4c8-152">3.5 (1996)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-153">' 97 (8.0)</span><span class="sxs-lookup"><span data-stu-id="de4c8-153">‘97 (8.0)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-154">5.0</span><span class="sxs-lookup"><span data-stu-id="de4c8-154">5.0</span></span></p></td>
<td><p><span data-ttu-id="de4c8-155">' 97 (8.0)</span><span class="sxs-lookup"><span data-stu-id="de4c8-155">‘97 (8.0)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-156">5.0</span><span class="sxs-lookup"><span data-stu-id="de4c8-156">5.0</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="de4c8-157">Microsoft Jet</span><span class="sxs-lookup"><span data-stu-id="de4c8-157">Microsoft Jet</span></span></p></td>
<td><p><span data-ttu-id="de4c8-158">4.0 (2000)</span><span class="sxs-lookup"><span data-stu-id="de4c8-158">4.0 (2000)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-159">2000 (9.0)</span><span class="sxs-lookup"><span data-stu-id="de4c8-159">2000 (9.0)</span></span></p></td>
<td><p></p></td>
<td><p><span data-ttu-id="de4c8-160">2000 (9.0)</span><span class="sxs-lookup"><span data-stu-id="de4c8-160">2000 (9.0)</span></span></p></td>
<td><p></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="de4c8-161">Moteur de base de données Microsoft Access</span><span class="sxs-lookup"><span data-stu-id="de4c8-161">Microsoft Access database engine</span></span></p></td>
<td><p><span data-ttu-id="de4c8-162">12.0 (2007)</span><span class="sxs-lookup"><span data-stu-id="de4c8-162">12.0 (2007)</span></span></p></td>
<td><p><span data-ttu-id="de4c8-163">2007</span><span class="sxs-lookup"><span data-stu-id="de4c8-163">2007</span></span></p></td>
<td><p></p></td>
<td><p></p></td>
<td><p></p></td>
</tr>
</tbody>
</table>
