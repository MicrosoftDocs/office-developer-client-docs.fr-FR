---
title: Description, propriété (ADO)
TOCTitle: Description Property (ADO)
ms:assetid: 31df5e36-641c-d213-31fc-6244e2983327
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249092(v=office.15)
ms:contentKeyID: 48544064
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 3837d60b58772bbe6b65af6673c4eaa471507e66
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469855"
---
# <a name="description-property-ado"></a><span data-ttu-id="4c86b-102">Description, propriété (ADO)</span><span class="sxs-lookup"><span data-stu-id="4c86b-102">Description Property (ADO)</span></span>


<span data-ttu-id="4c86b-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="4c86b-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="4c86b-104">Décrit un objet [Error](error-object-ado.md).</span><span class="sxs-lookup"><span data-stu-id="4c86b-104">Describes an [Error](error-object-ado.md) object.</span></span>

## <a name="return-value"></a><span data-ttu-id="4c86b-105">Valeur de retour</span><span class="sxs-lookup"><span data-stu-id="4c86b-105">Return Value</span></span>

<span data-ttu-id="4c86b-106">Renvoie une valeur de type **String** qui contient une description de l'erreur.</span><span class="sxs-lookup"><span data-stu-id="4c86b-106">Returns a **String** value that contains a description of the error.</span></span>

## <a name="remarks"></a><span data-ttu-id="4c86b-107">Notes</span><span class="sxs-lookup"><span data-stu-id="4c86b-107">Remarks</span></span>

<span data-ttu-id="4c86b-p101">Utilisez la propriété **Description** pour obtenir une courte description de l'erreur. Affichez cette propriété pour avertir l'utilisateur d'une erreur que vous ne pouvez pas ou ne souhaitez pas gérer. La chaîne provient d'ADO ou d'un fournisseur.</span><span class="sxs-lookup"><span data-stu-id="4c86b-p101">Use the **Description** property to obtain a short description of the error. Display this property to alert the user to an error that you cannot or do not want to handle. The string will come from either ADO or a provider.</span></span>

<span data-ttu-id="4c86b-p102">Les fournisseurs sont chargés de transmettre un texte d'erreur spécifique à ADO. ADO ajoute un objet [Error](error-object-ado.md) à la collection **Errors** pour chaque erreur ou avertissement de fournisseur reçu. Énumérez la collection **Errors** pour identifier les erreurs transmises par le fournisseur.</span><span class="sxs-lookup"><span data-stu-id="4c86b-p102">Providers are responsible for passing specific error text to ADO. ADO adds an [Error](error-object-ado.md) object to the **Errors** collection for each provider error or warning it receives. Enumerate the **Errors** collection to trace the errors that the provider passes.</span></span>
