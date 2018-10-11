---
title: Types d’événements (référence de base de données du bureau Access)
TOCTitle: Types of Events
ms:assetid: 94660fc1-65c3-1d21-c451-f3898014e0b6
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249660(v=office.15)
ms:contentKeyID: 48546414
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: ffb3bd48db4d8071e56553f4ea5bf79b83952466
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470754"
---
# <a name="types-of-events"></a><span data-ttu-id="96b8f-102">Types d'événements</span><span class="sxs-lookup"><span data-stu-id="96b8f-102">Types of Events</span></span>


<span data-ttu-id="96b8f-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="96b8f-103">**Applies to**: Access 2013 | Office 2013</span></span>



<span data-ttu-id="96b8f-p101">Il existe deux types d'événements de base. Les « événements Will », appelés avant le début d'une opération, comportent généralement le terme « Will » dans leur nom, par exemple, **WillChangeRecordset** ou **WillConnect**. Les événements appelés au terme d'un événement comportent généralement le terme « Complete » dans leur nom, par exemple, **RecordChangeComplete** ou **ConnectComplete**. Des exceptions existent, par exemple **InfoMessage**, mais ces événements se produisent une fois que l'opération associée est terminée.</span><span class="sxs-lookup"><span data-stu-id="96b8f-p101">There are two basic types of events. "Will Events," which are called before an operation starts, usually include "Will" in their names — for example, **WillChangeRecordset** or **WillConnect**. Events that are called after an event has been completed usually include "Complete" in their names — for example, **RecordChangeComplete** or **ConnectComplete**. Exceptions exist — such as **InfoMessage** — but these occur after the associated operation has completed.</span></span>

## <a name="will-events"></a><span data-ttu-id="96b8f-108">Événements Will</span><span class="sxs-lookup"><span data-stu-id="96b8f-108">Will Events</span></span>

<span data-ttu-id="96b8f-109">Les gestionnaires d'événements appelés avant le démarrage d'une opération vous permettent d'examiner ou de modifier les paramètres de l'opération, puis d'annuler l'opération ou de l'autoriser à se poursuivre.</span><span class="sxs-lookup"><span data-stu-id="96b8f-109">Event handlers called before the operation starts offer you the opportunity to examine or modify the operation parameters, and then either cancel the operation or allow it to complete.</span></span> <span data-ttu-id="96b8f-110">Ces routines de gestionnaire d’événements ont généralement des noms de la forme \* événement \*\*\*\*sera\*.</span><span class="sxs-lookup"><span data-stu-id="96b8f-110">These event-handler routines usually have names of the form \**Will*Event\*\*\*.</span></span>

## <a name="complete-events"></a><span data-ttu-id="96b8f-111">Événements Complete</span><span class="sxs-lookup"><span data-stu-id="96b8f-111">Complete Events</span></span>

<span data-ttu-id="96b8f-112">Les gestionnaires d'événements appelés au terme d'une opération peuvent avertir votre application qu'une opération est terminée.</span><span class="sxs-lookup"><span data-stu-id="96b8f-112">Event handlers called after an operation completes can notify your application that an operation has concluded.</span></span> <span data-ttu-id="96b8f-113">Ces gestionnaires reçoivent aussi une notification lorsqu'un gestionnaire d'événements Will annule une opération en attente.</span><span class="sxs-lookup"><span data-stu-id="96b8f-113">Such an event handler is also notified when a Will event handler cancels a pending operation.</span></span> <span data-ttu-id="96b8f-114">Ces routines de gestionnaire d’événements ont généralement des noms de la forme \***événement \* complète**.</span><span class="sxs-lookup"><span data-stu-id="96b8f-114">These event-handler routines usually have names of the form \***Event\*Complete**.</span></span>

<span data-ttu-id="96b8f-115">Les événements Will et Complete sont généralement utilisés par paire.</span><span class="sxs-lookup"><span data-stu-id="96b8f-115">Will and Complete events are typically used in pairs.</span></span>

## <a name="other-events"></a><span data-ttu-id="96b8f-116">Autres événements</span><span class="sxs-lookup"><span data-stu-id="96b8f-116">Other Events</span></span>

<span data-ttu-id="96b8f-117">Les autres gestionnaires d’événements — autrement dit, les événements dont les noms ne sont pas de la forme **sera \* événement**\* ou \***événement \* complète —** sont appelés qu’au terme d’une opération.</span><span class="sxs-lookup"><span data-stu-id="96b8f-117">The other event handlers — that is, events whose names are not of the form **Will\*Event**\* or \***Event\*Complete —** are called only after an operation completes.</span></span> <span data-ttu-id="96b8f-118">Il s'agit des événements **Disconnect**, **EndOfRecordset** et **InfoMessage**.</span><span class="sxs-lookup"><span data-stu-id="96b8f-118">These events are **Disconnect**, **EndOfRecordset**, and **InfoMessage**.</span></span>
