---
title: Espaces de noms (référence de base de données du bureau Access)
TOCTitle: Namespaces
ms:assetid: e39f003c-3d16-1fae-48c5-304593c41f2f
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250158(v=office.15)
ms:contentKeyID: 48548318
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: dc7c0cf302c0b8a297310dfa439ac87724331569
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471240"
---
# <a name="namespaces"></a><span data-ttu-id="06870-102">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="06870-102">Namespaces</span></span>


<span data-ttu-id="06870-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="06870-103">**Applies to**: Access 2013 | Office 2013</span></span>

## <a name="namespaces"></a><span data-ttu-id="06870-104">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="06870-104">Namespaces</span></span>

<span data-ttu-id="06870-105">Le format XML de persistance dans ADO utilise les quatre espaces de noms suivants.</span><span class="sxs-lookup"><span data-stu-id="06870-105">The XML persistence format in ADO uses the following four namespaces.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="06870-106">Préfixe</span><span class="sxs-lookup"><span data-stu-id="06870-106">Prefix</span></span></p></th>
<th><p><span data-ttu-id="06870-107">Description</span><span class="sxs-lookup"><span data-stu-id="06870-107">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="06870-108">s</span><span class="sxs-lookup"><span data-stu-id="06870-108">s</span></span></p></td>
<td><p><span data-ttu-id="06870-109">Fait référence à la &quot;XML-Data&quot; espace de noms contenant les éléments et attributs qui définissent le schéma du <strong>jeu d’enregistrements</strong>en cours.</span><span class="sxs-lookup"><span data-stu-id="06870-109">Refers to the &quot;XML-Data&quot; namespace containing the elements and attributes that define the schema of the current <strong>Recordset</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="06870-110">dt</span><span class="sxs-lookup"><span data-stu-id="06870-110">dt</span></span></p></td>
<td><p><span data-ttu-id="06870-111">Fait référence à la spécification des définitions du type de données.</span><span class="sxs-lookup"><span data-stu-id="06870-111">Refers to the data type definitions specification.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="06870-112">rs</span><span class="sxs-lookup"><span data-stu-id="06870-112">rs</span></span></p></td>
<td><p><span data-ttu-id="06870-113">Fait référence à l'espace de nom contenant les éléments et les attributs spécifiques aux propriétés et attributs d'un <strong>jeu d'enregistrements</strong> ADO.</span><span class="sxs-lookup"><span data-stu-id="06870-113">Refers to the namespace containing elements and attributes specific to ADO <strong>Recordset</strong> properties and attributes.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="06870-114">z</span><span class="sxs-lookup"><span data-stu-id="06870-114">z</span></span></p></td>
<td><p><span data-ttu-id="06870-115">Fait référence au schéma du jeu de lignes actif.</span><span class="sxs-lookup"><span data-stu-id="06870-115">Refers to the schema of the current rowset.</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="06870-p101">Il est déconseillé au client d'ajouter ses propres balises dans ces espaces de noms, comme défini dans les caractéristiques. Par exemple, il est déconseillé de définir un espace de noms comme « urn:schemas-microsoft-com:rowset » et d'écrire ensuite quelque chose comme « rs:MaPropreBalise ». Pour en savoir plus sur les espaces de noms, voir [Espaces de noms XML](https://www.w3.org/tr/xml-names/).</span><span class="sxs-lookup"><span data-stu-id="06870-p101">A client should not add its own tags to these namespaces, as defined by the specification. For example, a client should not define a namespace as "urn:schemas-microsoft-com:rowset" and then write out something such as "rs:MyOwnTag." To learn more about namespaces, see [XML Namespaces](https://www.w3.org/tr/xml-names/).</span></span>


> [!NOTE]
> <P><span data-ttu-id="06870-119">[!REMARQUE] L'ID de la balise du schéma doit être « RowsetSchema » et les espaces de noms utilisés pour faire référence au schéma du jeu de lignes actif doivent pointer vers « #RowsetSchema ».</span><span class="sxs-lookup"><span data-stu-id="06870-119">The ID for the schema tag must be "RowsetSchema," and the namespace used to refer to the schema of the current rowset must point to "#RowsetSchema."</span></span></P>



<span data-ttu-id="06870-120">Notez que le préfixe de l'espace de noms, c'est-à-dire la partie à droite des deux points et à gauche du signe égal, est arbitraire.</span><span class="sxs-lookup"><span data-stu-id="06870-120">Note that the prefix of the namespace, that part to the right of the colon and to the left of the equal sign, is arbitrary.</span></span>

```vb 
 
xmlns:rs="urn:schemas-microsoft-com:rowset" 
```

<span data-ttu-id="06870-p102">Vous pouvez lui attribuer n'importe quel nom à condition qu'il soit utilisé partout dans le document XML. ADO écrit toujours « s », « rs », « dt » et « z », mais ces préfixes ne sont pas pré-programmés dans le composant de chargement.</span><span class="sxs-lookup"><span data-stu-id="06870-p102">The user can define this to be any name as long as this name is consistently used throughout the XML document. ADO always writes out "s," "rs," "dt," and "z," but these prefix names are not hard-coded into the loading component.</span></span>

## <a name="namespaces"></a><span data-ttu-id="06870-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="06870-123">Namespaces</span></span>

<span data-ttu-id="06870-124">Le format XML de persistance dans ADO utilise les quatre espaces de noms suivants.</span><span class="sxs-lookup"><span data-stu-id="06870-124">The XML persistence format in ADO uses the following four namespaces.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="06870-125">Préfixe</span><span class="sxs-lookup"><span data-stu-id="06870-125">Prefix</span></span></p></th>
<th><p><span data-ttu-id="06870-126">Description</span><span class="sxs-lookup"><span data-stu-id="06870-126">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="06870-127">s</span><span class="sxs-lookup"><span data-stu-id="06870-127">s</span></span></p></td>
<td><p><span data-ttu-id="06870-128">Fait référence à la &quot;XML-Data&quot; espace de noms contenant les éléments et attributs qui définissent le schéma du <strong>jeu d’enregistrements</strong>en cours.</span><span class="sxs-lookup"><span data-stu-id="06870-128">Refers to the &quot;XML-Data&quot; namespace containing the elements and attributes that define the schema of the current <strong>Recordset</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="06870-129">dt</span><span class="sxs-lookup"><span data-stu-id="06870-129">dt</span></span></p></td>
<td><p><span data-ttu-id="06870-130">Fait référence à la spécification des définitions du type de données.</span><span class="sxs-lookup"><span data-stu-id="06870-130">Refers to the data type definitions specification.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="06870-131">rs</span><span class="sxs-lookup"><span data-stu-id="06870-131">rs</span></span></p></td>
<td><p><span data-ttu-id="06870-132">Fait référence à l'espace de nom contenant les éléments et les attributs spécifiques aux propriétés et attributs d'un <strong>jeu d'enregistrements</strong> ADO.</span><span class="sxs-lookup"><span data-stu-id="06870-132">Refers to the namespace containing elements and attributes specific to ADO <strong>Recordset</strong> properties and attributes.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="06870-133">z</span><span class="sxs-lookup"><span data-stu-id="06870-133">z</span></span></p></td>
<td><p><span data-ttu-id="06870-134">Fait référence au schéma du jeu de lignes actif.</span><span class="sxs-lookup"><span data-stu-id="06870-134">Refers to the schema of the current rowset.</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="06870-p103">Il est déconseillé au client d'ajouter ses propres balises dans ces espaces de noms, comme défini dans les caractéristiques. Par exemple, il est déconseillé de définir un espace de noms comme « urn:schemas-microsoft-com:rowset » et d'écrire ensuite quelque chose comme « rs:MaPropreBalise ». Pour en savoir plus sur les espaces de noms, voir [Espaces de noms XML](https://www.w3.org/tr/xml-names/).</span><span class="sxs-lookup"><span data-stu-id="06870-p103">A client should not add its own tags to these namespaces, as defined by the specification. For example, a client should not define a namespace as "urn:schemas-microsoft-com:rowset" and then write out something such as "rs:MyOwnTag." To learn more about namespaces, see [XML Namespaces](https://www.w3.org/tr/xml-names/).</span></span>


> [!NOTE]
> <P><span data-ttu-id="06870-138">[!REMARQUE] L'ID de la balise du schéma doit être « RowsetSchema » et les espaces de noms utilisés pour faire référence au schéma du jeu de lignes actif doivent pointer vers « #RowsetSchema ».</span><span class="sxs-lookup"><span data-stu-id="06870-138">The ID for the schema tag must be "RowsetSchema," and the namespace used to refer to the schema of the current rowset must point to "#RowsetSchema."</span></span></P>



<span data-ttu-id="06870-139">Notez que le préfixe de l'espace de noms, c'est-à-dire la partie à droite des deux points et à gauche du signe égal, est arbitraire.</span><span class="sxs-lookup"><span data-stu-id="06870-139">Note that the prefix of the namespace, that part to the right of the colon and to the left of the equal sign, is arbitrary.</span></span>

```vb 
 
xmlns:rs="urn:schemas-microsoft-com:rowset" 
```

<span data-ttu-id="06870-p104">Vous pouvez lui attribuer n'importe quel nom à condition qu'il soit utilisé partout dans le document XML. ADO écrit toujours « s », « rs », « dt » et « z », mais ces préfixes ne sont pas pré-programmés dans le composant de chargement.</span><span class="sxs-lookup"><span data-stu-id="06870-p104">The user can define this to be any name as long as this name is consistently used throughout the XML document. ADO always writes out "s," "rs," "dt," and "z," but these prefix names are not hard-coded into the loading component.</span></span>
