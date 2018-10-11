---
title: DBEngine.LoginTimeout Property (DAO)
TOCTitle: LoginTimeout Property
ms:assetid: 81d14153-79c5-7860-b6a8-4079d2d7acf7
ms:mtpsurl: https://msdn.microsoft.com/library/Ff196648(v=office.15)
ms:contentKeyID: 48545964
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052923
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 53df33e3171cde0d861f738a852350ec20bf0dec
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470247"
---
# <a name="dbenginelogintimeout-property-dao"></a><span data-ttu-id="366c0-102">DBEngine.LoginTimeout Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="366c0-102">DBEngine.LoginTimeout Property (DAO)</span></span>


<span data-ttu-id="366c0-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="366c0-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="366c0-104">Définit ou renvoie le nombre de secondes devant s'écouler avant l'apparition d'une erreur lorsque vous essayez de vous connecter à une base de données ODBC.</span><span class="sxs-lookup"><span data-stu-id="366c0-104">Sets or returns the number of seconds before an error occurs when you attempt to log on to an ODBC database.</span></span>

## <a name="syntax"></a><span data-ttu-id="366c0-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="366c0-105">Syntax</span></span>

<span data-ttu-id="366c0-106">*expression* . LoginTimeout</span><span class="sxs-lookup"><span data-stu-id="366c0-106">*expression* .LoginTimeout</span></span>

<span data-ttu-id="366c0-107">*expression* Variable qui représente un objet **DBEngine** .</span><span class="sxs-lookup"><span data-stu-id="366c0-107">*expression* A variable that represents a **DBEngine** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="366c0-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="366c0-108">Remarks</span></span>

<span data-ttu-id="366c0-p101">La valeur de la propriété **LoginTimeout** est de 20 secondes par défaut. Si la propriété **LoginTimeout** a la valeur 0, il n'y a pas de délai imparti pour la connexion.</span><span class="sxs-lookup"><span data-stu-id="366c0-p101">The default **LoginTimeout** property setting is 20 seconds. When the **LoginTimeout** property is set to 0, no timeout occurs.</span></span>

<span data-ttu-id="366c0-p102">Si vous essayez de vous connecter à une base de données ODBC, comme Microsoft SQL Server, la connexion peut échouer suite à des erreurs réseau ou si le serveur ne fonctionne pas. Plutôt que d'attendre 20 secondes pour vous connecter, vous pouvez spécifier le temps d'attente avant l'apparition d'une erreur. La connexion au serveur s'effectue implicitement parmi une série d'événements, comme l'exécution d'une requête sur une base de données serveur externe.</span><span class="sxs-lookup"><span data-stu-id="366c0-p102">When you're attempting to log on to an ODBC database, such as Microsoft SQL Server, the connection can fail as a result of network errors or because the server isn't running. Rather than waiting for the default 20 seconds to connect, you can specify how long to wait before raising an error. Logging on to the server happens implicitly as part of a number of different events, such as running a query on an external server database.</span></span>
