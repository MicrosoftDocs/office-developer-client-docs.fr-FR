---
title: Data Shaping (référence de base de données du bureau Access)
TOCTitle: Data Shaping
ms:assetid: 650571cc-6874-2cdb-dd76-0804d1cc4e38
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249390(v=office.15)
ms:contentKeyID: 48545305
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: e498bc461c2c267dee741a2bd9a4a83eacaca935
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472359"
---
# <a name="data-shaping"></a><span data-ttu-id="8badc-102">Mise en forme des données</span><span class="sxs-lookup"><span data-stu-id="8badc-102">Data Shaping</span></span>


<span data-ttu-id="8badc-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="8badc-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="8badc-104">La mise en forme des données permet de définir les colonnes d'un objet **Recordset** mis en forme, les relations entre les entités représentées par les colonnes et le mode de remplissage de l'objet **Recordset** avec des données.</span><span class="sxs-lookup"><span data-stu-id="8badc-104">Data shaping enables you to define the columns of a shaped **Recordset**, the relationships between the entities represented by the columns, and the manner in which the **Recordset** is populated with data.</span></span>

<span data-ttu-id="8badc-105">Les colonnes d'un objet **Recordset** mis en forme peuvent contenir des données provenant d'un fournisseur de données tel que Microsoft SQL Server, des références à un autre objet **Recordset**, des valeurs calculées d'après une seule ligne d'un objet **Recordset**, des valeurs dérivées d'une opération portant sur une colonne entière d'un objet **Recordset**, ou ces colonnes peuvent encore être des colonnes vides, récemment créées.</span><span class="sxs-lookup"><span data-stu-id="8badc-105">Columns of a shaped **Recordset** may contain data from a data provider such as Microsoft SQL Server, references to another **Recordset**, values derived from a calculation on a single row of a **Recordset**, values derived from an operation over a column of an entire **Recordset**; or they may be a newly fabricated, empty column.</span></span>

<span data-ttu-id="8badc-106">Lorsque vous récupérez la valeur d’une colonne qui contient une référence à un autre **objet Recordset**, ADO retourne automatiquement le réel **Recordset** représenté par la référence.</span><span class="sxs-lookup"><span data-stu-id="8badc-106">When you retrieve the value of a column that contains a reference to another **Recordset**, ADO automatically returns the actual **Recordset** represented by the reference.</span></span> <span data-ttu-id="8badc-107">Un **jeu d’enregistrements** contenant un autre **jeu d’enregistrements** est appelé un jeu d’enregistrements hiérarchique.</span><span class="sxs-lookup"><span data-stu-id="8badc-107">A **Recordset** that contains another **Recordset** is called a hierarchical recordset.</span></span> <span data-ttu-id="8badc-108">Jeux d’enregistrements hiérarchiques présente une relation *parent-enfant* dans laquelle le *parent* est l’objet recordset contenant et l' *enfant* est le jeu contenu.</span><span class="sxs-lookup"><span data-stu-id="8badc-108">Hierarchical recordsets exhibit a *parent-child* relationship, in which the *parent* is the containing recordset and the *child* is the contained recordset.</span></span> <span data-ttu-id="8badc-109">La référence à un **objet Recordset** est en fait une référence à un sous-ensemble de l’enfant, appelé un *chapitre*.</span><span class="sxs-lookup"><span data-stu-id="8badc-109">The reference to a **Recordset** is actually a reference to a subset of the child, called a *chapter*.</span></span> <span data-ttu-id="8badc-110">**Plus d’un objet Recordset**peut faire référence à un seul parent.</span><span class="sxs-lookup"><span data-stu-id="8badc-110">A single parent may reference more than one child **Recordset**.</span></span>

<span data-ttu-id="8badc-p102">La syntaxe de commandes de mise en forme permet de créer par programme un objet **Recordset** mis en forme. Vous pouvez ensuite accéder aux composants de l'objet **Recordset** par programme ou via un contrôle visuel approprié. Une commande de mise en forme est émise comme tout autre texte de commande ADO.</span><span class="sxs-lookup"><span data-stu-id="8badc-p102">The shape command syntax enables you to programmatically create a shaped **Recordset**. You can then access the components of the **Recordset** programmatically or through an appropriate visual control. A shape command is issued like any other ADO command text.</span></span>

<span data-ttu-id="8badc-p103">Vous pouvez créer des objets **Recordset** hiérarchiques de deux manières différentes grâce à la syntaxe de commandes de mise en forme. La première consiste à ajouter un objet **Recordset** enfant à l'objet **Recordset** parent. Généralement, le parent et l'enfant possèdent en commun au moins un colonne : la valeur de la colonne dans une ligne du parent est identique à celle de la colonne dans toutes les lignes au niveau de l'enfant.</span><span class="sxs-lookup"><span data-stu-id="8badc-p103">You can make hierarchical **Recordset** objects in two ways with the shape command syntax. The first appends a child **Recordset** to a parent **Recordset**. The parent and child typically have at least one column in common: the value of the column in a row of the parent is the same as the value of the column in all rows of the child.</span></span>

<span data-ttu-id="8badc-117">La deuxième méthode génère un **jeu d’enregistrements** parent à partir d’un **objet Recordset**enfant.</span><span class="sxs-lookup"><span data-stu-id="8badc-117">The second way generates a parent **Recordset** from a child **Recordset**.</span></span> <span data-ttu-id="8badc-118">Les enregistrements dans **l’objet Recordset** enfant sont regroupés, généralement à l’aide de la clause BY et une ligne est ajoutée à **l’objet Recordset** parent pour chaque groupe résultant dans l’enfant.</span><span class="sxs-lookup"><span data-stu-id="8badc-118">The records in the child **Recordset** are grouped, typically using the BY clause, and one row is added to the parent **Recordset** for each resulting group in the child.</span></span> <span data-ttu-id="8badc-119">Si la clause BY est omise, **l’objet Recordset** enfant forment un seul groupe et **l’objet Recordset** parent contient exactement une ligne.</span><span class="sxs-lookup"><span data-stu-id="8badc-119">If the BY clause is omitted, the child **Recordset** will form a single group and the parent **Recordset** will contain exactly one row.</span></span> <span data-ttu-id="8badc-120">Cela est utile pour calculer des agrégats de « totaux » sur l’intégralité de **l’objet Recordset**enfant.</span><span class="sxs-lookup"><span data-stu-id="8badc-120">This is useful for computing "grand total" aggregates over the entire child **Recordset**.</span></span>

<span data-ttu-id="8badc-121">Quelle que soit la façon dont le parent de **que jeu d’enregistrements** est créé, il contient une colonne de chapitre qui sert à associer à un **objet Recordset**enfant.</span><span class="sxs-lookup"><span data-stu-id="8badc-121">Regardless of which way the parent **Recordset** is formed, it will contain a chapter column that is used to relate it to a child **Recordset**.</span></span> <span data-ttu-id="8badc-122">Si vous le souhaitez, **l’objet Recordset** parent peut également posséder des colonnes qui contiennent des fonctions d’agrégation (somme, MIN, MAX et ainsi de suite) sur les lignes enfants.</span><span class="sxs-lookup"><span data-stu-id="8badc-122">If you wish, the parent **Recordset** may also have columns that contain aggregates (SUM, MIN, MAX, and so on) over the child rows.</span></span> <span data-ttu-id="8badc-123">Le parent et **l’objet Recordset** enfant peut-être colonnes contenant une expression sur la ligne dans le **jeu d’enregistrements**, ainsi que les colonnes qui sont nouvelles et initialement vides.</span><span class="sxs-lookup"><span data-stu-id="8badc-123">Both the parent and the child **Recordset** may have columns which contain an expression on the row in the **Recordset**, as well as columns which are new and initially empty.</span></span>

<span data-ttu-id="8badc-124">Vous pouvez imbriquer des objets **Recordset** hiérarchiques à n'importe quelle profondeur (en d'autres termes, créer des objets **Recordset** enfants d'autres objets **Recordset** enfants, etc.).</span><span class="sxs-lookup"><span data-stu-id="8badc-124">You can nest hierarchical **Recordset** objects to any depth (that is, create child **Recordset** objects of child **Recordset** objects, and so on).</span></span>

<span data-ttu-id="8badc-125">Vous pouvez accéder aux composants **Recordset** de l'objet **Recordset** mis en forme par programme ou via un contrôle visuel adéquat.</span><span class="sxs-lookup"><span data-stu-id="8badc-125">You can access the **Recordset** components of the shaped **Recordset** programmatically or through an appropriate visual control.</span></span>

<span data-ttu-id="8badc-126">Pour obtenir des exemples de commandes de mise en forme et des hiérarchies qui en résultent, voir « Utilisation de Microsoft Data Shaping Service pour OLE DB ».</span><span class="sxs-lookup"><span data-stu-id="8badc-126">For examples of shape commands and their resulting hierarchies, see Using the Data Shaping Service for OLE DB: A Closer Look.</span></span>
