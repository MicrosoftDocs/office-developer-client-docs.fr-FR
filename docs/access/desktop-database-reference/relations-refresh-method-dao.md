---
title: Relations.Refresh Method (DAO)
TOCTitle: Refresh Method
ms:assetid: d71cecf2-da90-5f62-9e51-f994e660ad34
ms:mtpsurl: https://msdn.microsoft.com/library/Ff835058(v=office.15)
ms:contentKeyID: 48547997
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 7015f6e5475251e72604d50176debfbe20528dcb
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470875"
---
# <a name="relationsrefresh-method-dao"></a><span data-ttu-id="1c158-102">Relations.Refresh Method (DAO)</span><span class="sxs-lookup"><span data-stu-id="1c158-102">Relations.Refresh Method (DAO)</span></span>


<span data-ttu-id="1c158-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="1c158-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="1c158-104">Met à jour les objets de la collection spécifiée en fonction du schéma actuel de la base de données.</span><span class="sxs-lookup"><span data-stu-id="1c158-104">Updates the objects in the specified colletion to reflect the database's current schema.</span></span>

## <a name="syntax"></a><span data-ttu-id="1c158-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="1c158-105">Syntax</span></span>

<span data-ttu-id="1c158-106">*expression* . Actualiser</span><span class="sxs-lookup"><span data-stu-id="1c158-106">*expression* .Refresh</span></span>

<span data-ttu-id="1c158-107">*expression* Variable qui représente un objet **Relations** .</span><span class="sxs-lookup"><span data-stu-id="1c158-107">*expression* A variable that represents a **Relations** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="1c158-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="1c158-108">Remarks</span></span>

<span data-ttu-id="1c158-p101">La méthode **Refresh** s'avère utile dans les environnements multi-utilisateur où d'autres utilisateurs peuvent modifier la base de données. Vous pouvez également avoir besoin de l'appliquer aux collections indirectement affectées par les modifications apportées à la base de données. Par exemple, si vous modifiez une collection **Users**, vous devrez peut-être actualiser une collection **Groups** avant d'utiliser la collection **Groups**.</span><span class="sxs-lookup"><span data-stu-id="1c158-p101">Use the **Refresh** method in multiuser environments in which other users may change the database. You may also need to use it on any collections that are indirectly affected by changes to the database. For example, if you change a **Users** collection, you may need to refresh a **Groups** collection before using the **Groups** collection.</span></span>

<span data-ttu-id="1c158-p102">Une collection est remplie d'objets lors de son premier référencement et ne reflète pas automatiquement les modifications que les autres utilisateurs apportent par la suite. S'il est probable qu'un autre utilisateur a modifié une collection, appliquez immédiatement la méthode Refresh à la collection avant d'effectuer une tâche dans votre application qui suppose la présence ou l'absence d'un objet déterminé dans la collection. Vous serez ainsi assuré que la collection est parfaitement à jour. D'un autre côté, l'utilisation de la méthode Refresh peut inutilement ralentir les performances.</span><span class="sxs-lookup"><span data-stu-id="1c158-p102">A collection is filled with objects the first time it's referred to and won't automatically reflect subsequent changes other users make. If it's likely that another user has changed a collection, use the Refresh method on the collection immediately before carrying out any task in your application that assumes the presence or absence of a particular object in the collection. This will ensure that the collection is as up-to-date as possible. On the other hand, using Refresh can unnecessarily slow performance.</span></span>
