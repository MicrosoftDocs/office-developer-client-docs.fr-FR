---
title: InfoMessage, événement (ADO)
TOCTitle: InfoMessage Event (ADO)
ms:assetid: 5d4f487f-96c8-4cf6-60ab-583510d3096f
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249328(v=office.15)
ms:contentKeyID: 48545109
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 6cc3b69eb3310d8e717605edd3d6fae32d635806
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472480"
---
# <a name="infomessage-event-ado"></a><span data-ttu-id="38278-102">InfoMessage, événement (ADO)</span><span class="sxs-lookup"><span data-stu-id="38278-102">InfoMessage Event (ADO)</span></span>


<span data-ttu-id="38278-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="38278-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="38278-104">L'événement **InfoMessage** est appelé à chaque avertissement généré lors d'une opération **ConnectionEvent**.</span><span class="sxs-lookup"><span data-stu-id="38278-104">The **InfoMessage** event is called whenever a warning occurs during a **ConnectionEvent** operation.</span></span>

## <a name="syntax"></a><span data-ttu-id="38278-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="38278-105">Syntax</span></span>

<span data-ttu-id="38278-106">InfoMessage*pError*, *adStatus*, *pConnection*</span><span class="sxs-lookup"><span data-stu-id="38278-106">InfoMessage*pError*, *adStatus*, *pConnection*</span></span>

## <a name="parameters"></a><span data-ttu-id="38278-107">Paramètres</span><span class="sxs-lookup"><span data-stu-id="38278-107">Parameters</span></span>

  - <span data-ttu-id="38278-108">*pError*</span><span class="sxs-lookup"><span data-stu-id="38278-108">*pError*</span></span>

  - <span data-ttu-id="38278-p101">Objet [Error](error-object-ado.md). Ce paramètre contient toutes les erreurs générées lors de l'opération. Si plusieurs erreurs sont retournées, énumérez la collection **Errors** pour les identifier.</span><span class="sxs-lookup"><span data-stu-id="38278-p101">An [Error](error-object-ado.md) object. This parameter contains any errors that are returned. If multiple errors are returned, enumerate the **Errors** collection to find them.</span></span>

  - <span data-ttu-id="38278-112">*adStatus*</span><span class="sxs-lookup"><span data-stu-id="38278-112">*adStatus*</span></span>

  - [<span data-ttu-id="38278-113">EventStatusEnum</span><span class="sxs-lookup"><span data-stu-id="38278-113">EventStatusEnum</span></span>](eventstatusenum.md)
    
    <span data-ttu-id="38278-114">Avant que cet événement soit retourné, affectez la valeur **adStatusUnwantedEvent** à ce paramètre pour éviter toute notification ultérieure.</span><span class="sxs-lookup"><span data-stu-id="38278-114">Before this event returns, set this parameter to **adStatusUnwantedEvent** to prevent subsequent notifications.</span></span>

  - <span data-ttu-id="38278-115">*pConnection*</span><span class="sxs-lookup"><span data-stu-id="38278-115">*pConnection*</span></span>

  - <span data-ttu-id="38278-p102">Objet [Connection](connection-object-ado.md). Connexion pour laquelle l'avertissement a été émis. Par exemple, des avertissements peuvent être générés lors de l'ouverture d'un objet **Connection** ou pendant l'exécution d'un objet [Command](command-object-ado.md) sur un objet **Connection**.</span><span class="sxs-lookup"><span data-stu-id="38278-p102">A [Connection](connection-object-ado.md) object. The connection for which the warning occurred. For example, warnings can occur when opening a **Connection** object or executing a [Command](command-object-ado.md) on a **Connection**.</span></span>
