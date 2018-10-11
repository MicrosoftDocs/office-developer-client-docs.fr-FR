---
title: Mode immédiat (référence de base de données du bureau Access)
TOCTitle: Immediate Mode
ms:assetid: 61bd3645-6e84-2e3a-7814-37d8c1247df0
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249362(v=office.15)
ms:contentKeyID: 48545220
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 202626b656c79d703d6e06fc4311c9f3bef9fd93
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469477"
---
# <a name="immediate-mode"></a><span data-ttu-id="27e71-102">Mode Immédiat</span><span class="sxs-lookup"><span data-stu-id="27e71-102">Immediate Mode</span></span>


<span data-ttu-id="27e71-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="27e71-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="27e71-p101">Le mode Immédiat est activé lorsque la propriété **LockType** est définie sur **adLockOptimistic** ou **adLockPessimistic**. Dans ce mode, les modifications apportées à un enregistrement sont répercutées dans la source de données dès que vous déclarez que le travail sur une ligne est terminé en invoquant la méthode **Update**.</span><span class="sxs-lookup"><span data-stu-id="27e71-p101">Immediate mode is in effect when the **LockType** property is set to **adLockOptimistic** or **adLockPessimistic**. In immediate mode, changes to a record are propagated to the data source as soon as you declare the work on a row complete by calling the **Update** method.</span></span>

## <a name="calling-update"></a><span data-ttu-id="27e71-106">Invocation de la méthode Update</span><span class="sxs-lookup"><span data-stu-id="27e71-106">Calling Update</span></span>

<span data-ttu-id="27e71-107">Si vous quittez l'enregistrement ajouté ou modifié avant d'appeler la méthode **Update**, ADO appelle automatiquement la méthode **Update** pour enregistrer les modifications.</span><span class="sxs-lookup"><span data-stu-id="27e71-107">If you move from the record you are adding or editing before calling the **Update** method, ADO will automatically call **Update** to save the changes.</span></span> <span data-ttu-id="27e71-108">Vous devez appeler la méthode **CancelUpdate** avant la navigation si vous souhaitez annuler les modifications apportées à l’enregistrement actif ou ignorer un enregistrement nouvellement ajouté.</span><span class="sxs-lookup"><span data-stu-id="27e71-108">You must call the **CancelUpdate** method before navigation if you want to cancel any changes made to the current record or discard a newly added record.</span></span>

<span data-ttu-id="27e71-109">L'enregistrement actif reste actif après avoir invoqué la méthode **Update**.</span><span class="sxs-lookup"><span data-stu-id="27e71-109">The current record remains current after you call the **Update** method.</span></span>

## <a name="cancelupdate"></a><span data-ttu-id="27e71-110">CancelUpdate</span><span class="sxs-lookup"><span data-stu-id="27e71-110">CancelUpdate</span></span>

<span data-ttu-id="27e71-p103">Utilisez la méthode **CancelUpdate** pour annuler toute modification apportée à la ligne active ou ignorer une ligne ajoutée. Vous ne pouvez plus annuler les modifications de la ligne active ou d'une nouvelle ligne après avoir invoqué la méthode **Update**, sauf si ces modifications font partie d'une transaction que vous pouvez annuler avec la méthode **RollbackTrans** ou d'une mise à jour en lots. Dans ce dernier cas, vous pouvez annuler la **mise à jour** avec la méthode **CancelUpdate** ou **CancelBatch**.</span><span class="sxs-lookup"><span data-stu-id="27e71-p103">Use the **CancelUpdate** method to cancel any changes made to the current row or to discard a newly added row. You cannot cancel changes to the current row or a new row after you call the **Update** method, unless the changes are either part of a transaction that you can roll back with the **RollbackTrans** method or part of a batch update. In the case of a batch update, you can cancel the **Update** with the **CancelUpdate** or **CancelBatch** method.</span></span>

<span data-ttu-id="27e71-114">Si vous ajoutez une ligne au moment de l'invocation de la méthode **CancelUpdate**, la ligne active devient la ligne qui était active avant l'invocation de **AddNew**.</span><span class="sxs-lookup"><span data-stu-id="27e71-114">If you are adding a new row when you call the **CancelUpdate** method, the current row becomes the row that was current before the **AddNew** call.</span></span>

<span data-ttu-id="27e71-115">Si vous n'avez pas modifié la ligne active ou ajouté de ligne, l'invocation de la méthode **CancelUpdate** génère une erreur.</span><span class="sxs-lookup"><span data-stu-id="27e71-115">If you have not changed the current row or added a new row, calling the **CancelUpdate** method generates an error.</span></span>
