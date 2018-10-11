---
title: ActualiserEnregistrement, action de macro
TOCTitle: RefreshRecord Macro Action
ms:assetid: 68c90d7d-f59c-9e83-bc30-8f37cf5a3696
ms:mtpsurl: https://msdn.microsoft.com/library/Ff195261(v=office.15)
ms:contentKeyID: 48545396
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm62122
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 91c49ca4d453418a02d55163a023e853fadd5773
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470564"
---
# <a name="refreshrecord-macro-action"></a><span data-ttu-id="e8d45-102">ActualiserEnregistrement, action de macro</span><span class="sxs-lookup"><span data-stu-id="e8d45-102">RefreshRecord Macro Action</span></span>


<span data-ttu-id="e8d45-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="e8d45-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="e8d45-104">Vous pouvez utiliser l'action **ActualiserEnregistrement** pour mettre à jour la source d'enregistrement sous-jacente de la feuille de données ou du formulaire actif de manière à refléter les modifications apportées aux enregistrements dans le jeu en cours.</span><span class="sxs-lookup"><span data-stu-id="e8d45-104">You can use the **RefreshRecord** action to update the underlying record source for the active form or datasheet to reflect changes made to the records in the current set.</span></span>

## <a name="remarks"></a><span data-ttu-id="e8d45-105">Notes</span><span class="sxs-lookup"><span data-stu-id="e8d45-105">Remarks</span></span>

<span data-ttu-id="e8d45-p101">L'action **ActualiserEnregistrement** affiche uniquement les modifications apportées aux enregistrements dans le jeu actif. Étant donné que l'action **ActualiserEnregistrement** n'actualise pas réellement la base de données, le jeu actuel n'inclut pas les enregistrements qui ont été ajoutés et n'exclut pas les enregistrements qui ont été supprimés depuis la dernière actualisation de la base de données ; il n'exclut pas non plus les enregistrements qui ne répondent plus aux critères de la requête ou du filtre. Pour actualiser la base de données, utilisez la méthode **[Actualiser](requery-macro-action.md)**. Lorsque la source d'enregistrement pour un formulaire est actualisée, le jeu d'enregistrements actuel reflète exactement toutes les données de la source d'enregistrement.</span><span class="sxs-lookup"><span data-stu-id="e8d45-p101">the **RefreshRecord** action shows only changes made to records in the current set. Because the **RefreshRecord** action does not actually requery the database, the current set will not include records that have been added or exclude records that have been deleted since the database was last requeried; Nor will it exclude records that no longer satisfy the criteria of the query or filter. To requery the database, use the **[Requery](requery-macro-action.md)** method. When the record source for a form is requeried, the current set of records will accurately reflect all data in the record source.</span></span>

<span data-ttu-id="e8d45-110">Le comportement de cette action de macro varie selon que vous l'appelez dans une base de données cliente ou une base de données Web.</span><span class="sxs-lookup"><span data-stu-id="e8d45-110">The behavior of this macro action depends on whether you are calling it in a client database or a web database.</span></span>

## <a name="client-database"></a><span data-ttu-id="e8d45-111">Base de données cliente</span><span class="sxs-lookup"><span data-stu-id="e8d45-111">Client database</span></span>

<span data-ttu-id="e8d45-p102">Dans une base de données cliente, vous pouvez utiliser l'action **ActualiserEnregistrement** pour mettre à jour la source d'enregistrement sous-jacente de la feuille de données ou du formulaire actif de façon à refléter les modifications apportées aux données dans le jeu actuel. Les modifications comprennent celles apportées par l'utilisateur actuel ou par d'autres utilisateurs dans un environnement multi-utilisateur. Elle équivaut à la méthode **[Actualiser](https://msdn.microsoft.com/library/ff836021\(v=office.15\))**.</span><span class="sxs-lookup"><span data-stu-id="e8d45-p102">In a client database, you can use the **RefreshRecord** action to update the underlying record source for the active form or datasheet to reflect changes made to the data in the current set. Changes include those made by the current user or by other users in a multiuser environment. It is equivalent to the **[Refresh](https://msdn.microsoft.com/library/ff836021\(v=office.15\))** method.</span></span>

<span data-ttu-id="e8d45-115">L'action de macro **ActualiserEnregistrement** effectue les opérations suivantes dans une base de données cliente :</span><span class="sxs-lookup"><span data-stu-id="e8d45-115">The **RefreshRecord** macro action does the following in a client database:</span></span>

1.  <span data-ttu-id="e8d45-p103">Elle met à jour la source d'enregistrement pour la feuille de données ou le formulaire actif de façon à refléter les modifications apportées aux lignes dans le jeu actuel. Pour les tables liées ODBC, elle extrait de la source de données les modifications apportées aux enregistrements dans le jeu actuel.</span><span class="sxs-lookup"><span data-stu-id="e8d45-p103">Updates the record source for the active form or datasheet to reflect the changes made to rows in the current set. For ODBC linked tables, retrieves changes to records in the current set from the data source.</span></span>

2.  <span data-ttu-id="e8d45-118">Elle met à jour le jeu actuel de façon à refléter les modifications apportées.</span><span class="sxs-lookup"><span data-stu-id="e8d45-118">Updates the current set to reflect the changes.</span></span> <span data-ttu-id="e8d45-119">Si une ligne dans la source d’enregistrement a été supprimée, il est modifié pour afficher \#Deleted.</span><span class="sxs-lookup"><span data-stu-id="e8d45-119">If a row in the record source has been deleted, it is changed to show \#Deleted.</span></span>

3.  <span data-ttu-id="e8d45-120">Actualise l’actif ou la feuille de données pour afficher les enregistrements modifiés et les \#suppression des enregistrements, dans le jeu actuel.</span><span class="sxs-lookup"><span data-stu-id="e8d45-120">Refreshes the active or datasheet to display any changed records and any \#Deleted records, in the current set.</span></span>

4.  <span data-ttu-id="e8d45-121">Elle actualise tous les sous-formulaires et sous-états dans la feuille de données ou le formulaire actif.</span><span class="sxs-lookup"><span data-stu-id="e8d45-121">Requeries any subforms and subreports on the active form or datasheet.</span></span>

## <a name="web-database"></a><span data-ttu-id="e8d45-122">Base de données Web</span><span class="sxs-lookup"><span data-stu-id="e8d45-122">Web database</span></span>

<span data-ttu-id="e8d45-p105">Dans une base de données Web, vous pouvez utiliser l'action **ActualiserEnregistrement** pour mettre à jour la source d'enregistrement sous-jacente de la feuille de données ou du formulaire actif de façon à refléter les modifications apportées aux enregistrements dans le jeu actuel. Les modifications comprennent celles apportées par l'utilisateur actuel ou par d'autres utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="e8d45-p105">In a web database, you can use the **RefreshRecord** action to update the underlying record source for the active form or datasheet to reflect changes made to the records in the current set. Changes include those made by the current user or other users.</span></span>

<span data-ttu-id="e8d45-125">L'action de macro **ActualiserEnregistrement** effectue les opérations suivantes dans une base de données Web :</span><span class="sxs-lookup"><span data-stu-id="e8d45-125">The **RefreshRecord** macro action does the following in a web database:</span></span>

1.  <span data-ttu-id="e8d45-p106">Elle extrait les modifications à partir du serveur pour les tables de base dans le jeu actuel. Pour les tables liées ODBC, elle extrait de la source de données les modifications apportées aux enregistrements dans le jeu actuel.</span><span class="sxs-lookup"><span data-stu-id="e8d45-p106">Retrieves changes from the server for any base tables in the current set. For ODBC linked tables, retrieves changes to records in the current set from the data source.</span></span>

2.  <span data-ttu-id="e8d45-128">Elle met à jour le jeu actuel de façon à refléter les modifications apportées.</span><span class="sxs-lookup"><span data-stu-id="e8d45-128">Updates the current set to reflect the changes.</span></span> <span data-ttu-id="e8d45-129">Si une ligne dans le jeu actuel a été supprimée, il est modifié pour afficher \#Deleted.</span><span class="sxs-lookup"><span data-stu-id="e8d45-129">If a row in the current set has been deleted, it is changed to show \#Deleted.</span></span>

3.  <span data-ttu-id="e8d45-130">Actualise l’actif formulaire ou feuille de données pour afficher les enregistrements modifiés et les \#suppression des enregistrements, dans le jeu actuel.</span><span class="sxs-lookup"><span data-stu-id="e8d45-130">Refreshes the active form or datasheet to display any changed records and any \#Deleted records, in the current set.</span></span>

4.  <span data-ttu-id="e8d45-131">Elle actualise tous les sous-formulaires et sous-états dans la feuille de données ou le formulaire actif.</span><span class="sxs-lookup"><span data-stu-id="e8d45-131">Requeries any subforms and subreports on the active form or datasheet.</span></span>
