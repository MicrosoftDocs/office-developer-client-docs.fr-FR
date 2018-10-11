---
title: Objets de données ActiveX (ADO) événements
TOCTitle: ADO Events
ms:assetid: 84ca9525-99cb-4ba6-2a4d-172414b8f0cc
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249576(v=office.15)
ms:contentKeyID: 48546041
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 5938ed58e3b0074f69a4380fd4a0ac0be857084b
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472142"
---
# <a name="ado-events"></a><span data-ttu-id="c7cfe-102">Événements ADO</span><span class="sxs-lookup"><span data-stu-id="c7cfe-102">ADO Events</span></span>


<span data-ttu-id="c7cfe-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="c7cfe-103">**Applies to**: Access 2013 | Office 2013</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-104"><a href="begintranscomplete-committranscomplete-and-rollbacktranscomplete-events-ado.md">BeginTransComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-104"><a href="begintranscomplete-committranscomplete-and-rollbacktranscomplete-events-ado.md">BeginTransComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-105">Appelé après l'opération <strong>BeginTrans</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-105">Called after the <strong>BeginTrans</strong> operation.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-106"><a href="begintranscomplete-committranscomplete-and-rollbacktranscomplete-events-ado.md">CommitTransComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-106"><a href="begintranscomplete-committranscomplete-and-rollbacktranscomplete-events-ado.md">CommitTransComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-107">Appelé après l'opération <strong>CommitTrans</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-107">Called after the <strong>CommitTrans</strong> operation.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-108"><a href="connectcomplete-and-disconnect-events-ado.md">ConnectComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-108"><a href="connectcomplete-and-disconnect-events-ado.md">ConnectComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-109">Appelé après l'établissement d'une connexion.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-109">Called after a connection starts.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-110"><a href="connectcomplete-and-disconnect-events-ado.md">Se déconnecter</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-110"><a href="connectcomplete-and-disconnect-events-ado.md">Disconnect</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-111">Appelé après la fermeture d'une connexion.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-111">Called after a connection ends.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-112"><a href="endofrecordset-event-ado.md">EndOfRecordset</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-112"><a href="endofrecordset-event-ado.md">EndOfRecordset</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-113">Appelé lors d'une tentative de déplacement d'une ligne au-delà de la fin de l'objet <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-113">Called when there is an attempt to move to a row past the end of the <strong>Recordset</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-114"><a href="executecomplete-event-ado.md">ExecuteComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-114"><a href="executecomplete-event-ado.md">ExecuteComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-115">Appelé à la fin de l'exécution d'une commande.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-115">Called after a command has finished executing.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-116"><a href="fetchcomplete-event-ado.md">FetchComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-116"><a href="fetchcomplete-event-ado.md">FetchComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-117">Appelé après que tous les enregistrements d'une opération asynchrone de longue durée ont été récupérés dans l'objet <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-117">Called after all the records in a lengthy asynchronous operation have been retrieved into the <strong>Recordset</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-118"><a href="fetchprogress-event-ado.md">FetchProgress</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-118"><a href="fetchprogress-event-ado.md">FetchProgress</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-119">Appelé régulièrement au cours d'une opération asynchrone de longue durée pour indiquer le nombre de lignes récupérées jusque là dans l'objet <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-119">Called periodically during a lengthy asynchronous operation to report how many rows have currently been retrieved into the <strong>Recordset</strong>.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-120"><a href="willchangefield-and-fieldchangecomplete-events-ado.md">FieldChangeComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-120"><a href="willchangefield-and-fieldchangecomplete-events-ado.md">FieldChangeComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-121">Appelé après que la valeur d'un ou plusieurs objets <strong>Field</strong> a changé.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-121">Called after the value of one or more <strong>Field</strong> objects has changed.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-122"><a href="infomessage-event-ado.md">InfoMessage</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-122"><a href="infomessage-event-ado.md">InfoMessage</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-123">Appelé lorsqu'un avertissement se produit au cours d'une opération <strong>ConnectionEvent</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-123">Called whenever a warning occurs during a <strong>ConnectionEvent</strong> operation.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-124"><a href="willmove-and-movecomplete-events-ado.md">MoveComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-124"><a href="willmove-and-movecomplete-events-ado.md">MoveComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-125">Appelé après la modification de la position actuelle dans l'objet <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-125">Called after the current position in the <strong>Recordset</strong> changes.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-126"><a href="willchangerecord-and-recordchangecomplete-events-ado.md">RecordChangeComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-126"><a href="willchangerecord-and-recordchangecomplete-events-ado.md">RecordChangeComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-127">Appelé après la modification d'un ou de plusieurs enregistrements.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-127">Called after one or more records change.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-128"><a href="willchangerecordset-and-recordsetchangecomplete-events-ado.md">RecordsetChangeComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-128"><a href="willchangerecordset-and-recordsetchangecomplete-events-ado.md">RecordsetChangeComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-129">Appelé après la modification de l'objet <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-129">Called after the <strong>Recordset</strong> has changed.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-130"><a href="begintranscomplete-committranscomplete-and-rollbacktranscomplete-events-ado.md">RollbackTransComplete</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-130"><a href="begintranscomplete-committranscomplete-and-rollbacktranscomplete-events-ado.md">RollbackTransComplete</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-131">Appelé après l'opération <strong>RollbackTrans</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-131">Called after the <strong>RollbackTrans</strong> operation.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-132"><a href="willchangefield-and-fieldchangecomplete-events-ado.md">WillChangeField</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-132"><a href="willchangefield-and-fieldchangecomplete-events-ado.md">WillChangeField</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-133">Appelé avant qu'une opération en attente change la valeur d'un ou plusieurs objets <strong>Field</strong> dans l'objet <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-133">Called before a pending operation changes the value of one or more <strong>Field</strong> objects in the <strong>Recordset</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-134"><a href="willchangerecord-and-recordchangecomplete-events-ado.md">WillChangeRecord</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-134"><a href="willchangerecord-and-recordchangecomplete-events-ado.md">WillChangeRecord</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-135">Appelé avant la modification d'un ou plusieurs enregistrements (lignes) dans l'objet <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-135">Called before one or more records (rows) in the <strong>Recordset</strong> change.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-136"><a href="willchangerecordset-and-recordsetchangecomplete-events-ado.md">WillChangeRecordset</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-136"><a href="willchangerecordset-and-recordsetchangecomplete-events-ado.md">WillChangeRecordset</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-137">Appelé avant qu'une opération en attente modifie l'objet <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-137">Called before a pending operation changes the <strong>Recordset</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-138"><a href="willconnect-event-ado.md">WillConnect</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-138"><a href="willconnect-event-ado.md">WillConnect</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-139">Appelé avant le début d'une connexion.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-139">Called before a connection starts.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c7cfe-140"><a href="willexecute-event-ado.md">WillExecute</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-140"><a href="willexecute-event-ado.md">WillExecute</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-141">Appelé juste avant qu'une commande en attente ne s'exécute sur cette connexion pour permettre à l'utilisateur d'examiner et de modifier les paramètres de la commande.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-141">Called just before a pending command executes on this connection and affords the user an opportunity to examine and modify the pending execution parameters.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c7cfe-142"><a href="willmove-and-movecomplete-events-ado.md">WillMove</a></span><span class="sxs-lookup"><span data-stu-id="c7cfe-142"><a href="willmove-and-movecomplete-events-ado.md">WillMove</a></span></span></p></td>
<td><p><span data-ttu-id="c7cfe-143">L’événement <strong>WillMove</strong> est appelé <em>avant</em> une opération en attente modifie la position actuelle dans le <strong>jeu d’enregistrements</strong>.</span><span class="sxs-lookup"><span data-stu-id="c7cfe-143">The <strong>WillMove</strong> event is called <em>before</em> a pending operation changes the current position in the <strong>Recordset</strong>.</span></span></p></td>
</tr>
</tbody>
</table>
