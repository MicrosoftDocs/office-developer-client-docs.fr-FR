---
title: Personnalisation de l'objet DataFactory
TOCTitle: DataFactory Customization
ms:assetid: 43cd7416-1f05-87ee-22f0-6cf0d2d1b39f
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249205(v=office.15)
ms:contentKeyID: 48544511
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: b805f9d6ff7a1f3b27bb5600d42cabf8fce651cd
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471771"
---
# <a name="datafactory-customization"></a><span data-ttu-id="90a2d-102">Personnalisation de l'objet DataFactory</span><span class="sxs-lookup"><span data-stu-id="90a2d-102">DataFactory Customization</span></span>


<span data-ttu-id="90a2d-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="90a2d-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="90a2d-p101">RDS (Remote Data Service) permet d'accéder facilement aux données dans un système client/serveur à trois niveaux. Un contrôle des données client spécifie des paramètres de chaîne de connexion et de commande pour exécuter une requête sur une source de données distante ou des paramètres de chaîne de connexion et d'objet [Recordset](recordset-object-ado.md) pour effectuer des mises à jour.</span><span class="sxs-lookup"><span data-stu-id="90a2d-p101">Remote Data Service (RDS) provides a way to easily perform data access in a three-tier client/server system. A client data control specifies connection and command string parameters to perform a query on a remote data source, or connection string and [Recordset](recordset-object-ado.md) object parameters to perform an update.</span></span>

<span data-ttu-id="90a2d-p102">Les paramètres sont transmis à un programme serveur qui effectue l'opération d'accès aux données sur la source de données distante. RDS fournit un programme serveur par défaut, l'objet [RDSServer.DataFactory](datafactory-object-rdsserver.md). L'objet **RDSServer.DataFactory** retourne tout objet **Recordset** résultant d'une requête au client.</span><span class="sxs-lookup"><span data-stu-id="90a2d-p102">The parameters are passed to a server program, which performs the data-access operation on the remote data source. RDS provides a default server program called the [RDSServer.DataFactory](datafactory-object-rdsserver.md) object. The **RDSServer.DataFactory** object returns any **Recordset** object produced by a query to the client.</span></span>

<span data-ttu-id="90a2d-p103">Toutefois, l'objet **RDSServer.DataFactory** ne peut effectuer que des requêtes et des mises à jour. Il ne peut pas effectuer d'opérations de validation ni de traitement sur les chaînes de connexion ou de commande.</span><span class="sxs-lookup"><span data-stu-id="90a2d-p103">However, the **RDSServer.DataFactory** is limited to performing queries and updates. It cannot perform any validation or processing on the connection or command strings.</span></span>

<span data-ttu-id="90a2d-p104">Avec ADO, vous pouvez spécifier que **DataFactory** doit fonctionner en association avec un autre type de programme serveur appelé *gestionnaire*. Le gestionnaire peut modifier les chaînes de commande et de connexion cliente avant qu'elles soient utilisées pour accéder à la source de données. En outre, le gestionnaire peut appliquer des droits d'accès qui permettent au client de lire et d'écrire des données dans la source de données.</span><span class="sxs-lookup"><span data-stu-id="90a2d-p104">With ADO, you can specify that the **DataFactory** work in conjunction with another type of server program called a *handler*. The handler can modify client connection and command strings before they are used to access the data source. In addition, the handler can enforce access rights, which govern the ability of the client to read and write data to the data source.</span></span>

<span data-ttu-id="90a2d-114">Les paramètres utilisés par le gestionnaire pour modifier les paramètres et les droits d'accès du client sont spécifiés dans les sections d'un fichier de personnalisation.</span><span class="sxs-lookup"><span data-stu-id="90a2d-114">The parameters the handler uses to modify client parameters and access rights are specified in sections of a customization file.</span></span>

<span data-ttu-id="90a2d-115">Consultez les rubriques suivantes pour plus d'informations sur la personnalisation de l'objet **DataFactory**:</span><span class="sxs-lookup"><span data-stu-id="90a2d-115">See the following topics for more information about customizing the **DataFactory** object:</span></span>

  - [<span data-ttu-id="90a2d-116">Présentation du fichier de personnalisation</span><span class="sxs-lookup"><span data-stu-id="90a2d-116">Understanding the Customization File</span></span>](understanding-the-customization-file.md)

  - [<span data-ttu-id="90a2d-117">Section Connect du fichier de personnalisation</span><span class="sxs-lookup"><span data-stu-id="90a2d-117">Customization File Connect Section</span></span>](customization-file-connect-section.md)

  - [<span data-ttu-id="90a2d-118">Section SQL du fichier de personnalisation</span><span class="sxs-lookup"><span data-stu-id="90a2d-118">Customization File SQL Section</span></span>](customization-file-sql-section.md)

  - [<span data-ttu-id="90a2d-119">Section UserList du fichier de personnalisation</span><span class="sxs-lookup"><span data-stu-id="90a2d-119">Customization File UserList Section</span></span>](customization-file-userlist-section.md)

  - [<span data-ttu-id="90a2d-120">Section Logs du fichier de personnalisation</span><span class="sxs-lookup"><span data-stu-id="90a2d-120">Customization File Logs Section</span></span>](customization-file-logs-section.md)

  - <span data-ttu-id="90a2d-121">[Paramètres clients requis](https://msdn.microsoft.com/library/ff836356\(v=office.15\))</span><span class="sxs-lookup"><span data-stu-id="90a2d-121">[Required Client Settings](https://msdn.microsoft.com/library/ff836356\(v=office.15\))</span></span>

  - <span data-ttu-id="90a2d-122">[Création d'un gestionnaire personnalisé](https://msdn.microsoft.com/library/jj249402\(v=office.15\))</span><span class="sxs-lookup"><span data-stu-id="90a2d-122">[Writing Your Own Customized Handler](https://msdn.microsoft.com/library/jj249402\(v=office.15\))</span></span>
