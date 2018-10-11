---
title: DefaultDatabase, propriété (ADO)
TOCTitle: DefaultDatabase Property (ADO)
ms:assetid: a35c5631-f9d9-e51f-950b-e52169830d94
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249757(v=office.15)
ms:contentKeyID: 48546784
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 136e6bc0d68aac987a331ce20b8fca36a72afa73
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471093"
---
# <a name="defaultdatabase-property-ado"></a><span data-ttu-id="2e9b5-102">DefaultDatabase, propriété (ADO)</span><span class="sxs-lookup"><span data-stu-id="2e9b5-102">DefaultDatabase Property (ADO)</span></span>


<span data-ttu-id="2e9b5-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="2e9b5-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="2e9b5-104">Indique la base de données par défaut d'un objet [Connection](connection-object-ado.md).</span><span class="sxs-lookup"><span data-stu-id="2e9b5-104">Indicates the default database for a [Connection](connection-object-ado.md) object.</span></span>

## <a name="settings-and-return-values"></a><span data-ttu-id="2e9b5-105">Paramètres et valeurs de retour</span><span class="sxs-lookup"><span data-stu-id="2e9b5-105">Settings and Return Values</span></span>

<span data-ttu-id="2e9b5-106">Définit ou renvoie une valeur de type **String**, qui correspond au nom d'une base de données disponible du fournisseur.</span><span class="sxs-lookup"><span data-stu-id="2e9b5-106">Sets or returns a **String** value that evaluates to the name of a database available from the provider.</span></span>

## <a name="remarks"></a><span data-ttu-id="2e9b5-107">Notes</span><span class="sxs-lookup"><span data-stu-id="2e9b5-107">Remarks</span></span>

<span data-ttu-id="2e9b5-108">Utilisez la propriété **DefaultDatabase** pour définir ou retourner le nom de la base de données par défaut d'un objet **Connection** spécifique.</span><span class="sxs-lookup"><span data-stu-id="2e9b5-108">Use the **DefaultDatabase** property to set or return the name of the default database on a specific **Connection** object.</span></span>

<span data-ttu-id="2e9b5-p101">S'il existe une base de données par défaut, les chaînes SQL peuvent utiliser une syntaxe non qualifiée pour accéder aux objets de cette base de données. Pour accéder aux objets dans une base de données autre que celle spécifiée dans la propriété **DefaultDatabase**, vous devez qualifier les noms d'objet avec le nom de la base de données souhaitée. Lors de la connexion, le fournisseur écrit les informations de la base de données par défaut dans la propriété **DefaultDatabase**. Certains fournisseurs n'autorisent qu'une seule base de données par connexion, auquel cas vous ne pouvez pas modifier la propriété **DefaultDatabase**.</span><span class="sxs-lookup"><span data-stu-id="2e9b5-p101">If there is a default database, SQL strings may use an unqualified syntax to access objects in that database. To access objects in a database other than the one specified in the **DefaultDatabase** property, you must qualify object names with the desired database name. Upon connection, the provider will write default database information to the **DefaultDatabase** property. Some providers allow only one database per connection, in which case you cannot change the **DefaultDatabase** property.</span></span>

<span data-ttu-id="2e9b5-113">Il se peut que certains fournisseurs et sources de données ne prennent pas en charge cette fonctionnalité et retournent une erreur ou une chaîne vide.</span><span class="sxs-lookup"><span data-stu-id="2e9b5-113">Some data sources and providers may not support this feature, and may return an error or an empty string.</span></span>

<span data-ttu-id="2e9b5-114">**L’utilisation du Service de données à distance** Cette propriété n’est pas disponible sur un objet de **connexion** côté client.</span><span class="sxs-lookup"><span data-stu-id="2e9b5-114">**Remote Data Service Usage**This property is not available on a client-side **Connection** object.</span></span>
