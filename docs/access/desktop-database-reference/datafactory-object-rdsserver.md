---
title: DataFactory, objet (RDSServer)
TOCTitle: DataFactory Object (RDSServer)
ms:assetid: 1de76cdd-34dc-8547-29aa-48ad6067bdea
ms:mtpsurl: https://msdn.microsoft.com/library/JJ248971(v=office.15)
ms:contentKeyID: 48543605
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: ba705854fc27a18fd0d3d7a7c6fd7acdefb3830f
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469350"
---
# <a name="datafactory-object-rdsserver"></a><span data-ttu-id="fe005-102">DataFactory, objet (RDSServer)</span><span class="sxs-lookup"><span data-stu-id="fe005-102">DataFactory Object (RDSServer)</span></span>


<span data-ttu-id="fe005-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="fe005-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="fe005-104">Cet objet métier côté serveur par défaut implémente des méthodes qui procurent un accès en lecture-écriture aux sources de données spécifiées pour des applications côté client.</span><span class="sxs-lookup"><span data-stu-id="fe005-104">This default server-side business object implements methods that provide read/write data access to specified data sources for client-side applications.</span></span>

## <a name="remarks"></a><span data-ttu-id="fe005-105">Notes</span><span class="sxs-lookup"><span data-stu-id="fe005-105">Remarks</span></span>

<span data-ttu-id="fe005-p101">L'objet **RDSServer.DataFactory** est conçu comme un objet Automation côté serveur, qui reçoit les demandes clientes. Dans une implémentation Internet, il réside sur un serveur Web et est instancié par le composant ADISAPI. L'objet **RDSServer.DataFactory** fournit un accès en lecture-écriture aux sources de données spécifiées, mais ne contient pas de logique de validation ou de règles métier.</span><span class="sxs-lookup"><span data-stu-id="fe005-p101">The **RDSServer.DataFactory** object is designed as a server-side Automation object that receives client requests. In an Internet implementation, it resides on a Web server and is instantiated by the ADISAPI component. The **RDSServer.DataFactory** object provides read and write access to specified data sources, but doesn't contain any validation or business rules logic.</span></span>

<span data-ttu-id="fe005-p102">Si vous utilisez une méthode disponible à la fois pour les objets **RDSServer.DataFactory** et [RDS.DataControl](datacontrol-object-rds.md), Remote Data Service utilise la version **RDS.DataControl** par défaut. Un scénario de programmation de base est pris comme point de départ par défaut, **RDSServer.DataFactory** étant utilisé comme objet métier côté serveur générique.</span><span class="sxs-lookup"><span data-stu-id="fe005-p102">If you use a method that is available in both the **RDSServer.DataFactory** and [RDS.DataControl](datacontrol-object-rds.md) objects, Remote Data Service uses the **RDS.DataControl** version by default. The default assumes a basic programming scenario, where the **RDSServer.DataFactory** serves as a generic server-side business object.</span></span>

<span data-ttu-id="fe005-111">Si vous voulez que votre application Web gère le traitement spécifique des tâches côté serveur, vous pouvez remplacer **RDSServer.DataFactory** par un objet métier personnalisé.</span><span class="sxs-lookup"><span data-stu-id="fe005-111">If you want your Web application to handle task-specific server-side processing, you can replace the **RDSServer.DataFactory** with a custom business object.</span></span>

<span data-ttu-id="fe005-p103">Vous pouvez créer des objets métier côté serveur qui appellent les méthodes de **RDSServer.DataFactory**, comme [Query](query-method-rds.md) et [CreateRecordset](createrecordset-method-rds.md). Cela s'avère utile si vous voulez ajouter des fonctionnalités à vos objets métier tout en profitant des technologies Remote Data Service existantes.</span><span class="sxs-lookup"><span data-stu-id="fe005-p103">You can create server-side business objects that call the **RDSServer.DataFactory** methods, such as [Query](query-method-rds.md) and [CreateRecordset](createrecordset-method-rds.md). This is helpful if you want to add functionality to your business objects, but take advantage of existing Remote Data Service technologies.</span></span>

<span data-ttu-id="fe005-114">L'identificateur de classe de l'objet **RDSServer.DataFactory** est 9381D8F5-0288-11D0-9501-00AA00B911A5.</span><span class="sxs-lookup"><span data-stu-id="fe005-114">The class ID for the **RDSServer.DataFactory** object is 9381D8F5-0288-11D0-9501-00AA00B911A5.</span></span>
