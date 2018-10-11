---
title: 'Annexe A : Fournisseurs'
TOCTitle: 'Appendix A: Providers'
ms:assetid: b3f92279-8d66-ad59-71c4-c0448168125a
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249857(v=office.15)
ms:contentKeyID: 48547207
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: dbd9536edd15f923af85f2fadad8b696077af4a4
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471611"
---
# <a name="appendix-a-providers"></a><span data-ttu-id="58fb0-102">Annexe A : Fournisseurs</span><span class="sxs-lookup"><span data-stu-id="58fb0-102">Appendix A: Providers</span></span>


<span data-ttu-id="58fb0-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="58fb0-103">**Applies to**: Access 2013 | Office 2013</span></span>


<span data-ttu-id="58fb0-p101">Cette section s'intéresse à trois types de fournisseurs : les fournisseurs de données, les fournisseurs de services et les composants de services. Il existe deux catégories de fournisseurs : les fournisseurs de données et les fournisseurs de services. Un *fournisseur de données* est propriétaire de ses données et les expose sous forme de tableaux dans votre application. Un *fournisseur de services* encapsule un service en produisant et en utilisant des données pour enrichir les fonctionnalités de vos applications ADO. Il peut entrer également dans la sous-catégorie des *composants de services* s'il doit fonctionner conjointement avec d'autres fournisseurs ou composants de services.</span><span class="sxs-lookup"><span data-stu-id="58fb0-p101">This section addresses three kinds of providers: data providers, service providers, and service components. Providers fall into two categories: those providing data and those providing services. A *data provider* owns its own data and exposes it in tabular form to your application. A *service provider* encapsulates a service by producing and consuming data, augmenting features in your ADO applications. A service provider may also be further defined as a *service component*, which must work in conjunction with other service providers or components.</span></span>

## <a name="data-providers"></a><span data-ttu-id="58fb0-109">Fournisseurs de données</span><span class="sxs-lookup"><span data-stu-id="58fb0-109">Data Providers</span></span>

<span data-ttu-id="58fb0-110">ADO est un outil puissant et flexible capable de se connecter à n'importe quel fournisseur de données d'un groupe de fournisseurs différents tout en continuant d'exposer le même modèle de programmation, indépendamment des fonctionnalités spécifiques du fournisseur choisi.</span><span class="sxs-lookup"><span data-stu-id="58fb0-110">ADO is powerful and flexible because it can connect to any of several different data providers and still expose the same programming model, regardless of the specific features of any given provider.</span></span>

<span data-ttu-id="58fb0-p102">Cependant, comme chaque fournisseur de données est unique, le mode d'interaction de votre application avec ADO variera légèrement en fonction du fournisseur de données. Les différences portent généralement sur trois grands points :</span><span class="sxs-lookup"><span data-stu-id="58fb0-p102">However, because each data provider is unique, how your application interacts with ADO will vary slightly by data provider. The differences usually fall into one of three categories:</span></span>

  - <span data-ttu-id="58fb0-113">Paramètres de connexion dans la propriété [ConnectionString](connectionstring-property-ado.md).</span><span class="sxs-lookup"><span data-stu-id="58fb0-113">Connection parameters in the [ConnectionString](connectionstring-property-ado.md) property.</span></span>

  - <span data-ttu-id="58fb0-114">Utilisation de l'objet [Command](command-object-ado.md).</span><span class="sxs-lookup"><span data-stu-id="58fb0-114">[Command](command-object-ado.md) object usage.</span></span>

  - <span data-ttu-id="58fb0-115">Comportement de l'objet [Recordset](recordset-object-ado.md) spécifique au fournisseur.</span><span class="sxs-lookup"><span data-stu-id="58fb0-115">Provider-specific [Recordset](recordset-object-ado.md) behavior.</span></span>

<span data-ttu-id="58fb0-116">Les détails concernant chaque fournisseur de données de Microsoft actuellement disponible sont répertoriés ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="58fb0-116">Details for each of the data providers currently available from Microsoft are listed as follows.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="58fb0-117">Champ d'action</span><span class="sxs-lookup"><span data-stu-id="58fb0-117">Area</span></span></p></th>
<th><p><span data-ttu-id="58fb0-118">Rubrique</span><span class="sxs-lookup"><span data-stu-id="58fb0-118">Topic</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="58fb0-119">Bases de données ODBC</span><span class="sxs-lookup"><span data-stu-id="58fb0-119">ODBC databases</span></span></p></td>
<td><p><span data-ttu-id="58fb0-120"><a href="microsoft-ole-db-provider-for-odbc.md">Fournisseur Microsoft OLE DB pour ODBC</a></span><span class="sxs-lookup"><span data-stu-id="58fb0-120"><a href="microsoft-ole-db-provider-for-odbc.md">Microsoft OLE DB Provider for ODBC</a></span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="58fb0-121">Service d'indexation Microsoft</span><span class="sxs-lookup"><span data-stu-id="58fb0-121">Microsoft Indexing Service</span></span></p></td>
<td><p><span data-ttu-id="58fb0-122"><a href="microsoft-ole-db-provider-for-microsoft-indexing-service.md">Fournisseur Microsoft OLE DB pour le service d'indexation Microsoft</a></span><span class="sxs-lookup"><span data-stu-id="58fb0-122"><a href="microsoft-ole-db-provider-for-microsoft-indexing-service.md">Microsoft OLE DB Provider for Microsoft Indexing Service</a></span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="58fb0-123">Service Microsoft Active Directory</span><span class="sxs-lookup"><span data-stu-id="58fb0-123">Microsoft Active Directory Service</span></span></p></td>
<td><p><span data-ttu-id="58fb0-124"><a href="microsoft-ole-db-provider-for-microsoft-active-directory-service.md">Fournisseur Microsoft OLE DB pour le service Microsoft Active Directory</a></span><span class="sxs-lookup"><span data-stu-id="58fb0-124"><a href="microsoft-ole-db-provider-for-microsoft-active-directory-service.md">Microsoft OLE DB Provider for Microsoft Active Directory Service</a></span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="58fb0-125">Bases de données Microsoft Jet</span><span class="sxs-lookup"><span data-stu-id="58fb0-125">Microsoft Jet databases</span></span></p></td>
<td><p><span data-ttu-id="58fb0-126"><a href="microsoft-ole-db-provider-for-microsoft-jet.md">Fournisseur OLE DB pour Microsoft Jet</a></span><span class="sxs-lookup"><span data-stu-id="58fb0-126"><a href="microsoft-ole-db-provider-for-microsoft-jet.md">OLE DB Provider for Microsoft Jet</a></span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="58fb0-127">Microsoft SQL Server</span><span class="sxs-lookup"><span data-stu-id="58fb0-127">Microsoft SQL Server</span></span></p></td>
<td><p><span data-ttu-id="58fb0-128"><a href="microsoft-ole-db-provider-for-sql-server.md">Fournisseur Microsoft OLE DB pour SQL Server</a></span><span class="sxs-lookup"><span data-stu-id="58fb0-128"><a href="microsoft-ole-db-provider-for-sql-server.md">Microsoft OLE DB Provider for SQL Server</a></span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="58fb0-129">Bases de données Oracle</span><span class="sxs-lookup"><span data-stu-id="58fb0-129">Oracle databases</span></span></p></td>
<td><p><span data-ttu-id="58fb0-130"><a href="microsoft-ole-db-provider-for-oracle.md">Fournisseur Microsoft OLE DB pour Oracle</a></span><span class="sxs-lookup"><span data-stu-id="58fb0-130"><a href="microsoft-ole-db-provider-for-oracle.md">Microsoft OLE DB Provider for Oracle</a></span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="58fb0-131">Publication Internet</span><span class="sxs-lookup"><span data-stu-id="58fb0-131">Internet Publishing</span></span></p></td>
<td><p><span data-ttu-id="58fb0-132"><a href="microsoft-ole-db-provider-for-internet-publishing.md">Fournisseur Microsoft OLE DB pour la publication Internet</a></span><span class="sxs-lookup"><span data-stu-id="58fb0-132"><a href="microsoft-ole-db-provider-for-internet-publishing.md">Microsoft OLE DB Provider for Internet Publishing</a></span></span></p></td>
</tr>
</tbody>
</table>


## <a name="provider-specific-dynamic-properties"></a><span data-ttu-id="58fb0-133">Propriétés dynamiques spécifiques à un fournisseur</span><span class="sxs-lookup"><span data-stu-id="58fb0-133">Provider-Specific Dynamic Properties</span></span>

<span data-ttu-id="58fb0-p103">Les collections [Properties](properties-collection-ado.md) des objets [Connection](connection-object-ado.md), [Command](command-object-ado.md) et [Recordset](recordset-object-ado.md) incluent des propriétés dynamiques spécifiques au fournisseur. Ces propriétés donnent des informations sur les fonctionnalités spécifiques au fournisseur qui viennent compléter les propriétés intégrées prises en charge par ADO.</span><span class="sxs-lookup"><span data-stu-id="58fb0-p103">The [Properties](properties-collection-ado.md) collections of the [Connection](connection-object-ado.md), [Command](command-object-ado.md), and [Recordset](recordset-object-ado.md) objects include dynamic properties specific to the provider. These properties provide information about functionality specific to the provider beyond the built-in properties that ADO supports.</span></span>

<span data-ttu-id="58fb0-p104">Après avoir établi la connexion et après avoir créé ces objets, utilisez la méthode [Refresh](refresh-method-ado.md) sur la collection **Properties** de l'objet pour obtenir les propriétés spécifiques au fournisseur. Consultez la documentation de fournisseur et le guide OLE DB Programmer's Reference (en anglais) pour obtenir des informations détaillées sur ces propriétés dynamiques.</span><span class="sxs-lookup"><span data-stu-id="58fb0-p104">After establishing the connection and creating these objects, use the [Refresh](refresh-method-ado.md) method on the object's **Properties** collection to obtain the provider-specific properties. Refer to the provider documentation and the OLE DB Programmer's Reference for detailed information about these dynamic properties.</span></span>

## <a name="service-providers"></a><span data-ttu-id="58fb0-138">Fournisseurs de services</span><span class="sxs-lookup"><span data-stu-id="58fb0-138">Service Providers</span></span>

<span data-ttu-id="58fb0-p105">Pour utiliser un fournisseur de services, vous devez fournir un mot clé. Vous devez aussi connaître les propriétés dynamiques spécifiques à chaque fournisseur de services. Les détails spécifiques à chaque fournisseur sont répertoriés pour tous les fournisseurs de services Microsoft disponibles :</span><span class="sxs-lookup"><span data-stu-id="58fb0-p105">To use a service provider, you must supply a keyword. You should also be aware of the provider-specific dynamic properties associated with each service provider. Provider-specific details are listed for each of the service providers currently available from Microsoft:</span></span>

  - [<span data-ttu-id="58fb0-142">Microsoft Data Shaping Service pour OLE DB</span><span class="sxs-lookup"><span data-stu-id="58fb0-142">Microsoft Data Shaping Service for OLE DB</span></span>](microsoft-data-shaping-service-for-ole-db-ado-service-provider.md)

  - [<span data-ttu-id="58fb0-143">Fournisseur de persistance Microsoft OLE DB</span><span class="sxs-lookup"><span data-stu-id="58fb0-143">Microsoft OLE DB Persistence Provider</span></span>](microsoft-ole-db-persistence-provider-ado-service-provider.md)

  - [<span data-ttu-id="58fb0-144">Fournisseur Microsoft OLE DB d'accès à distance</span><span class="sxs-lookup"><span data-stu-id="58fb0-144">Microsoft OLE DB Remoting Provider</span></span>](microsoft-ole-db-remoting-provider-ado-service-provider.md)

## <a name="service-components"></a><span data-ttu-id="58fb0-145">Composants de services</span><span class="sxs-lookup"><span data-stu-id="58fb0-145">Service Components</span></span>

<span data-ttu-id="58fb0-p106">Le composant de services [Service de curseur pour OLE DB](microsoft-cursor-service-for-ole-db-ado-service-component.md) complète les fonctions de prise en charge du curseur des fournisseurs de données. Il exige également l'utilisation d'un mot clé et possède des propriétés dynamiques.</span><span class="sxs-lookup"><span data-stu-id="58fb0-p106">The [Cursor Service for OLE DB](microsoft-cursor-service-for-ole-db-ado-service-component.md) service component supplements the cursor support functions of data providers. It also requires a keyword and has dynamic properties.</span></span>

<span data-ttu-id="58fb0-148">Pour plus d'informations sur les fournisseurs, consultez la documentation relative à Microsoft OLE DB dans le kit de développement logiciel (SDK) des composants Microsoft Data Access ou visitez le [Centre Accès aux Données et Stockage](https://msdn.microsoft.com/data/default.aspx).</span><span class="sxs-lookup"><span data-stu-id="58fb0-148">For more information about providers, see the documentation for Microsoft OLE DB in the Microsoft Data Access Components SDK or visit the [Data Platform Developer Center](https://msdn.microsoft.com/data/default.aspx).</span></span>

## <a name="provider-commands"></a><span data-ttu-id="58fb0-149">Commandes fournisseur</span><span class="sxs-lookup"><span data-stu-id="58fb0-149">Provider Commands</span></span>

<span data-ttu-id="58fb0-150">Pour chaque fournisseur répertorié ici, si vos applications permettent aux utilisateurs d’entrer des instructions SQL en tant que commandes fournisseur, vous devez toujours valider l’entrée d’utilisateur et être vigilant des attaques possibles à l’aide des instructions SQL potentiellement dangereuses, telles que, dans le cadre de la entrée utilisateur.</span><span class="sxs-lookup"><span data-stu-id="58fb0-150">For each provider listed here, if your applications allow users to enter SQL statements as the provider commands, you must always validate the user input and be vigilant of possible hacker attacks using potentially dangerous SQL statement, such as, , as part of the user input.</span></span>
