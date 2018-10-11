---
title: DBEngine.Version Property (DAO)
TOCTitle: Version Property
ms:assetid: b2807dc1-604f-4423-289a-ff38a3d9f31b
ms:mtpsurl: https://msdn.microsoft.com/library/Ff822024(v=office.15)
ms:contentKeyID: 48547171
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052986
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 3516a25623b6838286969c51f2d9346321f3326e
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469295"
---
# <a name="dbengineversion-property-dao"></a><span data-ttu-id="ab9f4-102">DBEngine.Version Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="ab9f4-102">DBEngine.Version Property (DAO)</span></span>


<span data-ttu-id="ab9f4-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="ab9f4-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="ab9f4-p101">Renvoie la version de DAO en cours d'utilisation. Valeur de type **String** en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="ab9f4-p101">Rreturns the version of DAO currently in use. Read-only **String**.</span></span>

## <a name="syntax"></a><span data-ttu-id="ab9f4-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ab9f4-106">Syntax</span></span>

<span data-ttu-id="ab9f4-107">*expression* . Version</span><span class="sxs-lookup"><span data-stu-id="ab9f4-107">*expression* .Version</span></span>

<span data-ttu-id="ab9f4-108">*expression* Variable qui représente un objet **DBEngine** .</span><span class="sxs-lookup"><span data-stu-id="ab9f4-108">*expression* A variable that represents a **DBEngine** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="ab9f4-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="ab9f4-109">Remarks</span></span>

<span data-ttu-id="ab9f4-p102">La valeur renvoyée est une chaîne (type String) correspondant à un numéro de version au format « principale.secondaire ». Par exemple, « 3.0 ». Le numéro de version du produit comprend le numéro de version principale (3), un point, puis le numéro de version secondaire (0).</span><span class="sxs-lookup"><span data-stu-id="ab9f4-p102">The return value is a String that evaluates to a version number in the form "major.minor". For example, "3.0". The product version number consists of the version number (3), a period, and the release number (0).</span></span>
