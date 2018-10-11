---
title: ReadyState, propriété (RDS)
TOCTitle: ReadyState Property (RDS)
ms:assetid: e7b62205-a604-ef43-2f5d-9b51b46d2b5a
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250175(v=office.15)
ms:contentKeyID: 48548412
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 804a98a6fa7c93617a5b6e165a2a012969c66913
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471177"
---
# <a name="readystate-property-rds"></a><span data-ttu-id="e55a4-102">ReadyState, propriété (RDS)</span><span class="sxs-lookup"><span data-stu-id="e55a4-102">ReadyState Property (RDS)</span></span>


<span data-ttu-id="e55a4-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="e55a4-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="e55a4-104">Indique la progression d'un [DataControl](datacontrol-object-rds.md) lorsqu'il récupère des données dans son objet [Recordset](recordset-object-ado.md).</span><span class="sxs-lookup"><span data-stu-id="e55a4-104">Indicates the progress of a [DataControl](datacontrol-object-rds.md) object as it retrieves data into its [Recordset](recordset-object-ado.md) object.</span></span>

## <a name="settings-and-return-values"></a><span data-ttu-id="e55a4-105">Paramètres et valeurs renvoyées</span><span class="sxs-lookup"><span data-stu-id="e55a4-105">Settings and Return Values</span></span>

<span data-ttu-id="e55a4-106">Définit ou renvoie l'une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="e55a4-106">Sets or returns one of the following values.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="e55a4-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="e55a4-107">Value</span></span></p></th>
<th><p><span data-ttu-id="e55a4-108">Description</span><span class="sxs-lookup"><span data-stu-id="e55a4-108">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="e55a4-109"><strong>(adcReadyStateLoaded)</strong></span><span class="sxs-lookup"><span data-stu-id="e55a4-109"><strong>adcReadyStateLoaded</strong></span></span></p></td>
<td><p><span data-ttu-id="e55a4-p101">La requête est encore en cours d'exécution et aucune ligne n'a été extraite. L'objet <strong>Recordset</strong> de <strong>DataControl</strong> n'est pas utilisable.</span><span class="sxs-lookup"><span data-stu-id="e55a4-p101">The current query is still executing and no rows have been fetched. The <strong>DataControl</strong> object's <strong>Recordset</strong> is not available for use.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="e55a4-112"><strong>adcReadyStateInteractive</strong></span><span class="sxs-lookup"><span data-stu-id="e55a4-112"><strong>adcReadyStateInteractive</strong></span></span></p></td>
<td><p><span data-ttu-id="e55a4-p102">Le premier jeu de lignes récupéré par la requête actuelle a été stocké dans le <strong>Recordset</strong> de l'objet <strong>DataControl</strong> et peut être utilisé. Les lignes restantes sont encore en cours d'extraction.</span><span class="sxs-lookup"><span data-stu-id="e55a4-p102">An initial set of rows retrieved by the current query has been stored in the <strong>DataControl</strong> object's <strong>Recordset</strong> and are available for use. The remaining rows are still being fetched.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="e55a4-115"><strong>adcReadyStateComplete</strong></span><span class="sxs-lookup"><span data-stu-id="e55a4-115"><strong>adcReadyStateComplete</strong></span></span></p></td>
<td><p><span data-ttu-id="e55a4-116">Toutes les lignes extraites par la requête en cours ont été stockées dans le <strong>Recordset</strong> de l'objet <strong>DataControl</strong> et peuvent être utilisées.
</span><span class="sxs-lookup"><span data-stu-id="e55a4-116">All rows retrieved by the current query have been stored in the <strong>DataControl</strong> object's <strong>Recordset</strong> and are available for use.</span></span> <span data-ttu-id="e55a4-117">Cet état est aussi défini lorsque l'opération a été interrompue en raison d'une erreur ou de la non-initialisation de l'objet <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="e55a4-117">This state will also exist if an operation aborted due to an error, or if the <strong>Recordset</strong> object is not initialized.</span></span></p></td>
</tr>
</tbody>
</table>



> [!NOTE]
> <P><span data-ttu-id="e55a4-p104">[!REMARQUE] Chaque fichier exécutable côté client utilisant ces constantes doit fournir les déclarations correspondantes. Vous pouvez couper et coller les déclarations de constante dont vous avez besoin dans le fichier Adcvbs.inc placé dans le dossier C:\Program Files\Common Files\System\MSADC.</span><span class="sxs-lookup"><span data-stu-id="e55a4-p104">Each client-side executable file that uses these constants must provide declarations for them. You can cut and paste the constant declarations you want from the file Adcvbs.inc, located in the C:\Program Files\Common Files\System\MSADC folder.</span></span></P>



## <a name="remarks"></a><span data-ttu-id="e55a4-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="e55a4-120">Remarks</span></span>

<span data-ttu-id="e55a4-p105">Utilisez l'événement [onreadystatechange](onreadystatechange-event-rds.md) pour surveiller les modifications de la propriété **ReadyState** lors d'une requête asynchrone. C'est plus efficace que de vérifier régulièrement la valeur de la propriété.</span><span class="sxs-lookup"><span data-stu-id="e55a4-p105">Use the [onReadyStateChange](onreadystatechange-event-rds.md) event to monitor changes in the **ReadyState** property during an asynchronous query operation. This is more efficient than periodically checking the value of the property.</span></span>

<span data-ttu-id="e55a4-123">Si une erreur se produit pendant une opération asynchrone, la propriété **ReadyState** devient **adcReadyStateComplete**, la propriété [State](state-property-ado.md) passe de **adStateExecuting** à **adStateClosed**et le **jeu d’enregistrements** [valeur de](value-property-ado.md) propriété de l’objet reste *Nothing*.</span><span class="sxs-lookup"><span data-stu-id="e55a4-123">If an error occurs during an asynchronous operation, the **ReadyState** property changes to **adcReadyStateComplete**, the [State](state-property-ado.md) property changes from **adStateExecuting** to **adStateClosed**, and the **Recordset** object [Value](value-property-ado.md) property remains *Nothing*.</span></span>
