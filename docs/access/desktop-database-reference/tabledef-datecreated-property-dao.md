---
title: TableDef.DateCreated Property (DAO)
TOCTitle: DateCreated Property
ms:assetid: fedd28e9-41a4-db7f-9ba9-6ada350d594a
ms:mtpsurl: https://msdn.microsoft.com/library/Ff837292(v=office.15)
ms:contentKeyID: 48548947
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 6ffad5b1bee5756c969f03a920f453f1d3abf9e3
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469750"
---
# <a name="tabledefdatecreated-property-dao"></a><span data-ttu-id="a8bd1-102">TableDef.DateCreated Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="a8bd1-102">TableDef.DateCreated Property (DAO)</span></span>


<span data-ttu-id="a8bd1-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="a8bd1-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="a8bd1-p101">Renvoie la date et l'heure auxquelles un objet a été créé (espaces de travail Microsoft Access uniquement). Valeur **Variant** en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="a8bd1-p101">Returns the date and time that an object was created (Microsoft Access workspaces only). Read-only **Variant**.</span></span>

## <a name="syntax"></a><span data-ttu-id="a8bd1-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="a8bd1-106">Syntax</span></span>

<span data-ttu-id="a8bd1-107">*expression* . DateCreated</span><span class="sxs-lookup"><span data-stu-id="a8bd1-107">*expression* .DateCreated</span></span>

<span data-ttu-id="a8bd1-108">*expression* Variable qui représente un objet **TableDef** .</span><span class="sxs-lookup"><span data-stu-id="a8bd1-108">*expression* A variable that represents a **TableDef** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="a8bd1-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="a8bd1-109">Remarks</span></span>

<span data-ttu-id="a8bd1-p102">**DateCreated** et **LastUpdated** renvoient la date et l'heure auxquelles l'objet a été créé ou mis à jour en dernier lieu. Dans un environnement multi-utilisateur, les utilisateurs doivent obtenir ces paramètres directement du serveur de fichiers afin d'éviter des incohérences entre les paramètres des propriétés DateCreated et LastUpdated.</span><span class="sxs-lookup"><span data-stu-id="a8bd1-p102">**DateCreated** and **LastUpdated** return the date and time that the object was created or last updated. In a multiuser environment, users should get these settings directly from the file server to avoid discrepancies in the DateCreated and LastUpdated property settings.</span></span>
