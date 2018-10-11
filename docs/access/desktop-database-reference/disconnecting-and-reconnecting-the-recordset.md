---
title: Déconnexion et reconnexion de l'objet Recordset
TOCTitle: Disconnecting and Reconnecting the Recordset
ms:assetid: d608d95d-9a4e-17a1-107a-b88b77f3774c
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250077(v=office.15)
ms:contentKeyID: 48547975
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 90435606bb0b3059f5769c12fe7cf3cac0c8f9f3
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472396"
---
# <a name="disconnecting-and-reconnecting-the-recordset"></a><span data-ttu-id="99195-102">Déconnexion et reconnexion de l'objet Recordset</span><span class="sxs-lookup"><span data-stu-id="99195-102">Disconnecting and Reconnecting the Recordset</span></span>


<span data-ttu-id="99195-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="99195-103">**Applies to**: Access 2013 | Office 2013</span></span>

## <a name="disconnecting-and-reconnecting-the-recordset"></a><span data-ttu-id="99195-104">Déconnexion et reconnexion de l'objet Recordset</span><span class="sxs-lookup"><span data-stu-id="99195-104">Disconnecting and Reconnecting the Recordset</span></span>

<span data-ttu-id="99195-p101">Une des fonctionnalités les plus puissantes d'ADO permet d'ouvrir un objet **Recordset** côté client à partir d'une source de données et de *déconnecter* cet objet **Recordset** de la source de données. Après la déconnexion de l'objet **Recordset**, la connexion à la source de données peut être fermée, libérant ainsi les ressources serveur utilisées pour la gérer. Vous pouvez continuer à consulter et modifier les données dans l'objet **Recordset** pendant sa déconnexion puis le reconnecter ultérieurement à la source de données et envoyer vos mises à jour en mode de mise à jour par lot.</span><span class="sxs-lookup"><span data-stu-id="99195-p101">One of the most powerful features found in ADO is the capability to open a client-side **Recordset** from a data source and then *disconnect* the **Recordset** from the data source. Once the **Recordset** has been disconnected, the connection to the data source can be closed, thereby releasing the resources on the server used to maintain it. You can continue to view and edit the data in the **Recordset** while it is disconnected and later reconnect to the data source and send your updates in batch mode.</span></span>

<span data-ttu-id="99195-p102">Pour déconnecter un objet **Recordset**, ouvrez-le avec un emplacement de curseur dont la valeur est **adUseClient** puis affectez à la propriété **ActiveConnection** la valeur *Nothing*. (Les utilisateurs de C++ doivent affecter la valeur NULL à **ActiveConnection** pour effectuer une déconnexion.)</span><span class="sxs-lookup"><span data-stu-id="99195-p102">To disconnect a **Recordset**, open it with a cursor location of **adUseClient**, and then set the **ActiveConnection** property equal to *Nothing*. (C++ users should set the **ActiveConnection** equal to NULL to disconnect.)</span></span>

<span data-ttu-id="99195-110">Nous aurons recours à un objet **Recordset** déconnecté plus loin dans ce chapitre, lorsque nous aborderons la persistance de l'objet **Recordset**. Il sera utilisé pour illustrer un scénario dans lequel les données de l'objet **Recordset** doivent être mises à la disposition d'une application lorsque l'ordinateur client n'est pas connecté à un réseau.</span><span class="sxs-lookup"><span data-stu-id="99195-110">We will use a disconnected **Recordset** later in this chapter when we discuss **Recordset** persistence to address a scenario in which we need to have the data in a **Recordset** available to an application while the client computer is not connected to a network.</span></span>
