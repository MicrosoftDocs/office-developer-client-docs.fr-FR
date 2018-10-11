---
title: RechercherEnregistrement, action de macro
TOCTitle: SearchForRecord Macro Action
ms:assetid: a3483c41-adb5-5923-55f4-1a3c4f60cb2f
ms:mtpsurl: https://msdn.microsoft.com/library/Ff821023(v=office.15)
ms:contentKeyID: 48546781
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm118713
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 21efa96338363646be826248e6c9ca0b951496c9
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469495"
---
# <a name="searchforrecord-macro-action"></a><span data-ttu-id="85611-102">RechercherEnregistrement, action de macro</span><span class="sxs-lookup"><span data-stu-id="85611-102">SearchForRecord Macro Action</span></span>


<span data-ttu-id="85611-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="85611-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="85611-104">L'action **RechercherEnregistrement** permet de rechercher un enregistrement spécifique dans une table, une requête, un formulaire ou un état.</span><span class="sxs-lookup"><span data-stu-id="85611-104">You can use the **SearchForRecord** action to search for a specific record in a table, query, form or report.</span></span>

## <a name="setting"></a><span data-ttu-id="85611-105">Valeur</span><span class="sxs-lookup"><span data-stu-id="85611-105">Setting</span></span>

<span data-ttu-id="85611-106">L'action **RechercherEnregistrement** accepte les arguments suivants.</span><span class="sxs-lookup"><span data-stu-id="85611-106">The **SearchForRecord** action has the following arguments.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="85611-107">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="85611-107">Action argument</span></span></p></th>
<th><p><span data-ttu-id="85611-108">Description</span><span class="sxs-lookup"><span data-stu-id="85611-108">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="85611-109"><strong>Type d'objet</strong></span><span class="sxs-lookup"><span data-stu-id="85611-109"><strong>Object Type</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-p101">Entrez ou sélectionnez le type d'objet de base de données dans lequel vous souhaitez effectuer votre recherche. Vous pouvez sélectionner <strong>Table</strong>, <strong>Requête</strong>, <strong>Formulaire</strong> ou <strong>État</strong>.  </span><span class="sxs-lookup"><span data-stu-id="85611-p101">Enter or select the type of database object that you are searching in. You can select <strong>Table</strong>, <strong>Query</strong>, <strong>Form</strong>, or <strong>Report</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="85611-112"><strong>Nom de l'objet</strong></span><span class="sxs-lookup"><span data-stu-id="85611-112"><strong>Object Name</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-p102">Entrez ou sélectionnez l'objet spécifique qui contient l'enregistrement à rechercher. La liste déroulante affiche tous les objets de base de données du type sélectionné pour l'argument <strong>Type d'objet</strong>.  </span><span class="sxs-lookup"><span data-stu-id="85611-p102">Enter or select the specific object that contains the record to search for. The drop-down list shows all database objects of the type you selected for the <strong>Object Type</strong> argument.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="85611-115"><strong>Enregistrement</strong></span><span class="sxs-lookup"><span data-stu-id="85611-115"><strong>Record</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-116">Spécifiez le point de départ et la direction de la recherche.</span><span class="sxs-lookup"><span data-stu-id="85611-116">Specify the starting point and direction of the search.</span></span></p>
<div class="tableSection">
<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="85611-117">Valeur</span><span class="sxs-lookup"><span data-stu-id="85611-117">Setting</span></span></p></th>
<th><p><span data-ttu-id="85611-118">Description</span><span class="sxs-lookup"><span data-stu-id="85611-118">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="85611-119"><strong>Précédent</strong></span><span class="sxs-lookup"><span data-stu-id="85611-119"><strong>Previous</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-120">Recherche en arrière à partir de l'enregistrement actif.</span><span class="sxs-lookup"><span data-stu-id="85611-120">Search backward from the current record.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="85611-121"><strong>Suivant</strong></span><span class="sxs-lookup"><span data-stu-id="85611-121"><strong>Next</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-122">Recherche vers l'avant à partir de l'enregistrement actif.</span><span class="sxs-lookup"><span data-stu-id="85611-122">Search forward from the current record.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="85611-123"><strong>Premier</strong></span><span class="sxs-lookup"><span data-stu-id="85611-123"><strong>First</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-p103">Recherche vers l'avant à partir du premier enregistrement. Il s'agit de la valeur par défaut pour cet argument.</span><span class="sxs-lookup"><span data-stu-id="85611-p103">Search forward from the first record. This is the default value for this argument.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="85611-126"><strong>Dernier</strong></span><span class="sxs-lookup"><span data-stu-id="85611-126"><strong>Last</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-127">Recherche en arrière à partir du dernier enregistrement.</span><span class="sxs-lookup"><span data-stu-id="85611-127">Search backward from the last record.</span></span></p></td>
</tr>
</tbody>
</table>

</div></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="85611-128"><strong>Where Condition</strong></span><span class="sxs-lookup"><span data-stu-id="85611-128"><strong>Where Condition</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-129">Entrez les critères de la recherche à l’aide de la même syntaxe comme une clause SQL WHERE, sans le mot &quot;où&quot;.</span><span class="sxs-lookup"><span data-stu-id="85611-129">Enter the criteria for the search using the same syntax as an SQL WHERE clause, only without the word &quot;WHERE&quot;.</span></span> <span data-ttu-id="85611-130">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="85611-130">For example,</span></span></p>
<p>`Description = "Beverages"`</p>
<p><span data-ttu-id="85611-131">Pour créer un critère qui comprend une valeur dans une zone de texte d'un formulaire, vous devez créer une expression qui concatène la première partie du critère avec le nom de la zone de texte contenant la valeur à rechercher.</span><span class="sxs-lookup"><span data-stu-id="85611-131">To create a criterion that includes a value from a text box on a form, you must create an expression that concatenates the first part of the criterion with the name of the text box containing the value for which to search.</span></span> <span data-ttu-id="85611-132">Par exemple, le critère suivant recherche le champ Description de la valeur dans la zone de texte txtDescription du formulaire frmCategories.</span><span class="sxs-lookup"><span data-stu-id="85611-132">For example, the following criterion will search the Description field for the value in the text box named txtDescription on the form named frmCategories.</span></span> <span data-ttu-id="85611-133">Notez le signe égal (<strong>=</strong>) au début de l’expression et l’utilisation de guillemets simples (<strong>'</strong>) de chaque côté de la référence de zone de texte :</span><span class="sxs-lookup"><span data-stu-id="85611-133">Note the equal sign (<strong>=</strong>) at the beginning of the expression, and the use of single quotation marks (<strong>'</strong>) on either side of the text box reference:</span></span></p>
<p>`="Description = ' " & Forms![frmCategories]![txtDescription] & "'"`</p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="85611-134">Notes</span><span class="sxs-lookup"><span data-stu-id="85611-134">Remarks</span></span>

- <span data-ttu-id="85611-135">Lorsque plusieurs enregistrements correspondent aux critères dans l'argument **Condition Where**, les facteurs suivants déterminent l'enregistrement trouvé :</span><span class="sxs-lookup"><span data-stu-id="85611-135">In cases where more than one record matches the criteria in the **Where Condition** argument, the following factors determine which record is found:</span></span>
    
  - <span data-ttu-id="85611-136">**La définition de l’argument Record**Voir la table dans la section Paramètres pour plus d’informations sur l’argument **Record**.</span><span class="sxs-lookup"><span data-stu-id="85611-136">**The Record argument setting**See the table in the Settings section for more information about the **Record** argument.</span></span>
    
  - <span data-ttu-id="85611-137">\*\*L’ordre de tri des enregistrements \*\*Par exemple, si l’argument **Record** est défini sur **First**, le fait de changer l’ordre de tri des enregistrements peut permettre de trouver un autre enregistrement.</span><span class="sxs-lookup"><span data-stu-id="85611-137">**The sort order of the records**For example, if the **Record** argument is set to **First**, changing the sort order of the records might change which record is found.</span></span>

- <span data-ttu-id="85611-p106">Vous devez ouvrir l'objet spécifié dans l'argument **Nom de l'objet** avant d'exécuter cette action. Sinon, une erreur se produit.</span><span class="sxs-lookup"><span data-stu-id="85611-p106">The object specified in the **Object Name** argument must be open before this action is run. Otherwise, an error occurs.</span></span>

- <span data-ttu-id="85611-140">Si les critères de l'argument **Condition Where** ne sont pas remplis, aucune erreur ne se produit et le focus reste sur l'enregistrement actif.</span><span class="sxs-lookup"><span data-stu-id="85611-140">If the criteria in the **Where Condition** argument are not met, no error occurs and the focus remains on the current record.</span></span>

- <span data-ttu-id="85611-p107">Si vous recherchez l'enregistrement suivant ou précédent, la recherche s'arrête lorsqu'elle atteint la fin des données. Si aucun autre enregistrement ne correspond aux critères, aucune erreur ne se produit et le focus reste sur l'enregistrement actif. Pour confirmer qu'un enregistrement a été trouvé, entrez une condition pour l'action suivante, en faisant en sorte que la condition soit la même que les critères de l'argument **Condition Where**.</span><span class="sxs-lookup"><span data-stu-id="85611-p107">When searching for the previous or next record, the search does not "wrap" when it reaches the end of the data. If there are no further records that match the criteria, no error occurs and the focus remains on the current record. To confirm that a match was found, you can enter a condition for the next action, and make the condition the same as the criteria in the **Where Condition** argument.</span></span>

- <span data-ttu-id="85611-144">Pour exécuter l'action **RechercherEnregistrement** dans un module VBA, utilisez la méthode **RechercherEnregistrement** de l'objet **DoCmd**.</span><span class="sxs-lookup"><span data-stu-id="85611-144">To run the **SearchForRecord** action in a VBA module, use the **SearchForRecord** method of the **DoCmd** object.</span></span>

- <span data-ttu-id="85611-p108">L'action **RechercherEnregistrement** est similaire à l'action **[TrouverEnregistrement](findrecord-macro-action.md)**, à la différence près que **RechercherEnregistrement** comporte des fonctionnalités de recherche plus puissantes. L'action **RechercherEnregistrement** est essentiellement destinée à la recherche de chaînes et elle reproduit la fonctionnalité de la boîte de dialogue **Rechercher**. L'action **RechercherEnregistrement** utilise des critères similaires à ceux d'un filtre ou d'une requête SQL. La liste ci-après illustre certaines opérations que vous pouvez effectuer à l'aide de l'action **RechercherEnregistrement**:</span><span class="sxs-lookup"><span data-stu-id="85611-p108">The **SearchForRecord** action is similar to the **[FindRecord](findrecord-macro-action.md)** action, but **SearchForRecord** has more powerful search features. The **FindRecord** action is primarily used for finding strings, and it duplicates the functionality of the **Find** dialog box. The **SearchForRecord** action uses criteria that are more like those of a filter or an SQL query. The following list demonstrates some things you can do with the **SearchForRecord** action:</span></span>
    
  - <span data-ttu-id="85611-149">Vous pouvez utiliser des critères complexes dans l'argument **Condition Where**, tels que :</span><span class="sxs-lookup"><span data-stu-id="85611-149">You can use complex criteria in the **Where Condition** argument, such as</span></span>
        
    `Description = "Beverages" and CategoryID = 11`
    
  - <span data-ttu-id="85611-p109">Vous pouvez faire référence à des champs qui se trouvent dans la source d’enregistrement d’un formulaire ou d’un état mais qui n’apparaissent pas dans le formulaire ou l’état. Dans l’exemple précédent, ni Description ni CategoryID ne doivent apparaître dans le formulaire ou l’état pour que les critères fonctionnent.</span><span class="sxs-lookup"><span data-stu-id="85611-p109">You can refer to fields that are in the record source of a form or report but aren't displayed on the form or report. In the preceding example, neither Description nor CategoryID must be displayed on the form or report for the criteria to work.</span></span>
    
  - <span data-ttu-id="85611-p110">Vous pouvez utiliser des opérateurs logiques, tels que **\<**, **\>**, **ET**, **OU** et **ENTRE**. L'action **TrouverEnregistrement** ne produit que des résultats qui sont égaux à, commencent par ou contiennent la chaîne recherchée.</span><span class="sxs-lookup"><span data-stu-id="85611-p110">You can use logical operators, such as **\<**, **\>**, **AND**, **OR**, and **BETWEEN**. The **FindRecord** action only matches strings that equal, start with, or contain the string being searched for.</span></span>

## <a name="example"></a><span data-ttu-id="85611-154">Exemple</span><span class="sxs-lookup"><span data-stu-id="85611-154">Example</span></span>

<span data-ttu-id="85611-p111">La macro suivante ouvre d'abord la table Catégories en utilisant l'action **OuvrirTable**. Elle utilise ensuite l'action **RechercherEnregistrement** pour rechercher le premier enregistrement dans la table pour lequel le champ Description correspond à « Boissons ».</span><span class="sxs-lookup"><span data-stu-id="85611-p111">The following macro first opens the Categories table by using the **OpenTable** action. The macro then uses the **SearchForRecord** action to find the first record in the table where the Description field equals "Beverages."</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="85611-157">Action</span><span class="sxs-lookup"><span data-stu-id="85611-157">Action</span></span></p></th>
<th><p><span data-ttu-id="85611-158">Arguments</span><span class="sxs-lookup"><span data-stu-id="85611-158">Arguments</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="85611-159"><strong>OuvrirTable</strong></span><span class="sxs-lookup"><span data-stu-id="85611-159"><strong>OpenTable</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-160"><strong>Nom de la table</strong>: catégories<strong>affichage</strong>: <strong>Mode DatasheetData</strong>: <strong>Modifier</strong></span><span class="sxs-lookup"><span data-stu-id="85611-160"><strong>Table Name</strong>: Categories<strong>View</strong>: <strong>DatasheetData Mode</strong>: <strong>Edit</strong></span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="85611-161"><strong>RechercherEnregistrement</strong></span><span class="sxs-lookup"><span data-stu-id="85611-161"><strong>SearchForRecord</strong></span></span></p></td>
<td><p><span data-ttu-id="85611-162"><strong>Type d’objet</strong>: <strong>TableObject nom</strong>: catégories<strong>enregistrement</strong>: <strong>Condition FirstWhere</strong>: Description = &quot;boissons&quot;</span><span class="sxs-lookup"><span data-stu-id="85611-162"><strong>Object Type</strong>: <strong>TableObject Name</strong>: Categories<strong>Record</strong>: <strong>FirstWhere Condition</strong>: Description = &quot;Beverages&quot;</span></span></p></td>
</tr>
</tbody>
</table>
