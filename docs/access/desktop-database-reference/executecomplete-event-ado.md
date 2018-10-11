---
title: ExecuteComplete, événement (ADO)
TOCTitle: ExecuteComplete Event (ADO)
ms:assetid: 47317d97-e373-32f4-9438-2dff46b8d367
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249219(v=office.15)
ms:contentKeyID: 48544589
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: c24b6c681c857dfa66f424d174bd7ce4a2cbd772
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469128"
---
# <a name="executecomplete-event-ado"></a><span data-ttu-id="9418a-102">ExecuteComplete, événement (ADO)</span><span class="sxs-lookup"><span data-stu-id="9418a-102">ExecuteComplete Event (ADO)</span></span>


<span data-ttu-id="9418a-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="9418a-103">**Applies to**: Access 2013 | Office 2013</span></span>



<span data-ttu-id="9418a-104">L'événement **ExecuteComplete** est appelé au terme de l'exécution d'une commande.</span><span class="sxs-lookup"><span data-stu-id="9418a-104">The **ExecuteComplete** event is called after a command has finished executing.</span></span>

## <a name="syntax"></a><span data-ttu-id="9418a-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="9418a-105">Syntax</span></span>

<span data-ttu-id="9418a-106">ExecuteComplete*RecordsAffected* *pError*, *adStatus*, *pCommand*, *Connection*, *pConnection*</span><span class="sxs-lookup"><span data-stu-id="9418a-106">ExecuteComplete*RecordsAffected*, *pError*, *adStatus*, *pCommand*, *pRecordset*, *pConnection*</span></span>

## <a name="parameters"></a><span data-ttu-id="9418a-107">Paramètres</span><span class="sxs-lookup"><span data-stu-id="9418a-107">Parameters</span></span>

  - <span data-ttu-id="9418a-108">*RecordsAffected*</span><span class="sxs-lookup"><span data-stu-id="9418a-108">*RecordsAffected*</span></span>

  - <span data-ttu-id="9418a-109">Valeur de type **Long** indiquant le nombre d'enregistrements affectés par la commande.</span><span class="sxs-lookup"><span data-stu-id="9418a-109">A **Long** value indicating the number of records affected by the command.</span></span>

  - <span data-ttu-id="9418a-110">*pError*</span><span class="sxs-lookup"><span data-stu-id="9418a-110">*pError*</span></span>

  - <span data-ttu-id="9418a-p101">Objet [Error](error-object-ado.md), décrivant l'erreur qui s'est produite si **adStatus** a la valeur **adStatusErrorsOccurred**. Dans le cas contraire, il n'est pas défini.</span><span class="sxs-lookup"><span data-stu-id="9418a-p101">An [Error](error-object-ado.md) object. It describes the error that occurred if the value of **adStatus** is **adStatusErrorsOccurred**; otherwise it is not set.</span></span>

  - <span data-ttu-id="9418a-113">*adStatus*</span><span class="sxs-lookup"><span data-stu-id="9418a-113">*adStatus*</span></span>

  - [<span data-ttu-id="9418a-114">EventStatusEnum</span><span class="sxs-lookup"><span data-stu-id="9418a-114">EventStatusEnum</span></span>](eventstatusenum.md)
    
    <span data-ttu-id="9418a-115">Avant que cet événement soit retourné, affectez la valeur **adStatusUnwantedEvent** à ce paramètre pour éviter toute notification ultérieure.</span><span class="sxs-lookup"><span data-stu-id="9418a-115">Before this event returns, set this parameter to **adStatusUnwantedEvent** to prevent subsequent notifications.</span></span>

  - <span data-ttu-id="9418a-116">*pCommand*</span><span class="sxs-lookup"><span data-stu-id="9418a-116">*pCommand*</span></span>

  - <span data-ttu-id="9418a-117">L’objet [Command](command-object-ado.md) qui a été exécutée.</span><span class="sxs-lookup"><span data-stu-id="9418a-117">The [Command](command-object-ado.md) object that was executed.</span></span> <span data-ttu-id="9418a-118">Contient un objet **Command** même lors de l’appel de **méthode Connection.Execute** ou **Recordset.Open** sans création explicite d’une **commande**dans les cas de l’objet **Command** est créée en interne par ADO.</span><span class="sxs-lookup"><span data-stu-id="9418a-118">Contains a **Command** object even when calling **Connection.Execute** or **Recordset.Open** without explicitly creating a **Command**, in which cases the **Command** object is created internally by ADO.</span></span>

  - <span data-ttu-id="9418a-119">*Connection*</span><span class="sxs-lookup"><span data-stu-id="9418a-119">*pRecordset*</span></span>

  - <span data-ttu-id="9418a-p103">Objet [Recordset](recordset-object-ado.md) résultant de la commande exécutée. Cet objet **Recordset** peut être vide. Il ne doit jamais être détruit à partir de ce gestionnaire d'événements. En effet, cela entraîne une violation d'accès puisqu'ADO tente d'accéder à un objet qui n'existe plus.</span><span class="sxs-lookup"><span data-stu-id="9418a-p103">A [Recordset](recordset-object-ado.md) object that is the result of the executed command. This **Recordset** may be empty. You should never destroy this Recordset object from within this event handler. Doing so will result in an Access Violation when ADO tries to access an object that no longer exists.</span></span>

  - <span data-ttu-id="9418a-124">*pConnection*</span><span class="sxs-lookup"><span data-stu-id="9418a-124">*pConnection*</span></span>

  - <span data-ttu-id="9418a-p104">Objet [Connection](connection-object-ado.md), représentant la connexion utilisée pour l'exécution de l'opération.</span><span class="sxs-lookup"><span data-stu-id="9418a-p104">A [Connection](connection-object-ado.md) object. The connection over which the operation was executed.</span></span>

## <a name="remarks"></a><span data-ttu-id="9418a-127">Notes</span><span class="sxs-lookup"><span data-stu-id="9418a-127">Remarks</span></span>

<span data-ttu-id="9418a-128">Un événement **ExecuteComplete** peut se produire avec les méthodes **Connection.**[Execute](https://msdn.microsoft.com/library/jj249832\(v=office.15\)), **Command.**[Execute](https://msdn.microsoft.com/library/jj248785\(v=office.15\)), **Recordset.**[Open](open-method-ado-recordset.md), **Recordset.**[Requery](requery-method-ado.md) ou **Recordset.**[NextRecordset](nextrecordset-method-ado.md).</span><span class="sxs-lookup"><span data-stu-id="9418a-128">An **ExecuteComplete** event may occur due to the **Connection.**[Execute](https://msdn.microsoft.com/library/jj249832\(v=office.15\)), **Command.**[Execute](https://msdn.microsoft.com/library/jj248785\(v=office.15\)), **Recordset.**[Open](open-method-ado-recordset.md), **Recordset.**[Requery](requery-method-ado.md), or **Recordset.**[NextRecordset](nextrecordset-method-ado.md) methods.</span></span>
