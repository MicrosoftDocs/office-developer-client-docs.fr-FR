---
title: CommandTimeout, propriété (ADO)
TOCTitle: CommandTimeout Property (ADO)
ms:assetid: a0b6209c-9feb-08ae-002a-15d1d20734a8
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249739(v=office.15)
ms:contentKeyID: 48546714
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- ado210.chm1231124
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 58d95f6a3238d09ae10ddb3ba127a9fa68631f4f
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469878"
---
# <a name="commandtimeout-property-ado"></a><span data-ttu-id="8fe7e-102">CommandTimeout, propriété (ADO)</span><span class="sxs-lookup"><span data-stu-id="8fe7e-102">CommandTimeout Property (ADO)</span></span>


<span data-ttu-id="8fe7e-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="8fe7e-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="8fe7e-104">Indique le délai d'attente pour l'exécution d'une commande avant l'abandon de la tentative et la génération d'une erreur.</span><span class="sxs-lookup"><span data-stu-id="8fe7e-104">Indicates how long to wait while executing a command before terminating the attempt and generating an error.</span></span>

## <a name="settings-and-return-values"></a><span data-ttu-id="8fe7e-105">Paramètres et valeurs de retour</span><span class="sxs-lookup"><span data-stu-id="8fe7e-105">Settings and Return Values</span></span>

<span data-ttu-id="8fe7e-p101">Définit ou renvoie une valeur de type **Long** indiquant, en secondes, le délai d'attente pour l'exécution d'une commande. La valeur par défaut est 30.</span><span class="sxs-lookup"><span data-stu-id="8fe7e-p101">Sets or returns a **Long** value that indicates, in seconds, how long to wait for a command to execute. Default is 30.</span></span>

## <a name="remarks"></a><span data-ttu-id="8fe7e-108">Notes</span><span class="sxs-lookup"><span data-stu-id="8fe7e-108">Remarks</span></span>

<span data-ttu-id="8fe7e-p102">Utilisez la propriété **CommandTimeout** dans un objet [Connection](connection-object-ado.md) ou [Command](command-object-ado.md) pour permettre l'annulation d'un appel de la méthode [Execute](https://msdn.microsoft.com/library/jj248785\(v=office.15\)), en raison d'un délai dû au trafic réseau ou à une utilisation intensive du serveur. Si l'intervalle défini dans la propriété **CommandTimeout** s'écoule avant que la commande ait fini de s'exécuter, une erreur se produit et ADO annule la commande. Si vous attribuez la valeur 0 à la propriété, ADO attend indéfiniment que l'exécution se termine. Assurez-vous que le fournisseur et la source de données dans lesquels vous écrivez le code prennent en charge la fonctionnalité **CommandTimeout**.</span><span class="sxs-lookup"><span data-stu-id="8fe7e-p102">Use the **CommandTimeout** property on a [Connection](connection-object-ado.md) object or [Command](command-object-ado.md) object to allow the cancellation of an [Execute](https://msdn.microsoft.com/library/jj248785\(v=office.15\)) method call, due to delays from network traffic or heavy server use. If the interval set in the **CommandTimeout** property elapses before the command completes execution, an error occurs and ADO cancels the command. If you set the property to zero, ADO will wait indefinitely until the execution is complete. Make sure the provider and data source to which you are writing code support the **CommandTimeout** functionality.</span></span>

<span data-ttu-id="8fe7e-113">Le paramètre de **CommandTimeout** d'un objet **Connection** n'a aucun effet sur le paramètre de **CommandTimeout** d'un objet **Command** sur le même objet **Connection**. En d'autres termes, la propriété **CommandTimeout** de l'objet **Command** n'hérite pas de la valeur de la propriété **CommandTimeout** de l'objet **Connection**.</span><span class="sxs-lookup"><span data-stu-id="8fe7e-113">The **CommandTimeout** setting on a **Connection** object has no effect on the **CommandTimeout** setting on a **Command** object on the same **Connection**; that is, the **Command** object's **CommandTimeout** property does not inherit the value of the **Connection** object's **CommandTimeout** value.</span></span>

<span data-ttu-id="8fe7e-114">Sur un objet **Connection** , la propriété **CommandTimeout** reste en lecture/écriture après l’ouverture de la **connexion** .</span><span class="sxs-lookup"><span data-stu-id="8fe7e-114">On a **Connection** object, the **CommandTimeout** property remains read/write after the **Connection** is opened.</span></span>
