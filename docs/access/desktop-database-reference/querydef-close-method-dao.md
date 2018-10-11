---
title: QueryDef.Close Method (DAO)
TOCTitle: Close Method
ms:assetid: b2b63462-453d-9e2b-0bb3-69a4a7a6ecef
ms:mtpsurl: https://msdn.microsoft.com/library/Ff822031(v=office.15)
ms:contentKeyID: 48547179
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052976
f1_categories:
- Office.Version=v15
ms.openlocfilehash: e17936286a4bd661d0a210c905cf0b63b70cf936
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470921"
---
# <a name="querydefclose-method-dao"></a><span data-ttu-id="24c87-102">QueryDef.Close Method (DAO)</span><span class="sxs-lookup"><span data-stu-id="24c87-102">QueryDef.Close Method (DAO)</span></span>


<span data-ttu-id="24c87-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="24c87-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="24c87-104">Ferme un objet **QueryDef** ouvert.</span><span class="sxs-lookup"><span data-stu-id="24c87-104">Closes an open **QueryDef**.</span></span>

## <a name="syntax"></a><span data-ttu-id="24c87-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="24c87-105">Syntax</span></span>

<span data-ttu-id="24c87-106">*expression* . Fermer</span><span class="sxs-lookup"><span data-stu-id="24c87-106">*expression* .Close</span></span>

<span data-ttu-id="24c87-107">*expression* Variable qui représente un objet **QueryDef** .</span><span class="sxs-lookup"><span data-stu-id="24c87-107">*expression* A variable that represents a **QueryDef** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="24c87-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="24c87-108">Remarks</span></span>

<span data-ttu-id="24c87-109">Si l'objet **QueryDef** est déjà fermé lorsque vous utilisez **Close**, une erreur d'exécution se produit.</span><span class="sxs-lookup"><span data-stu-id="24c87-109">If the **QueryDef** object is already closed when you use **Close**, a run-time error occurs.</span></span>

<span data-ttu-id="24c87-110">Une alternative à la méthode **Close** est la valeur d’une variable d’objet sur **Nothing** (Set dbsTemp = Nothing).</span><span class="sxs-lookup"><span data-stu-id="24c87-110">An alternative to the **Close** method is to set the value of an object variable to **Nothing** (Set dbsTemp = Nothing).</span></span>
