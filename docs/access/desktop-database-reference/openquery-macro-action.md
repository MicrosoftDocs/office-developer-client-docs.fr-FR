---
title: OuvrirRequête, action de macro
TOCTitle: OpenQuery Macro Action
ms:assetid: 64bfce73-aeaf-9a78-895c-492e5da43ded
ms:mtpsurl: https://msdn.microsoft.com/library/Ff195094(v=office.15)
ms:contentKeyID: 48545298
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm89069
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 30d4f245618e305be0e46b75cc0ea5c224444fa8
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471186"
---
# <a name="openquery-macro-action"></a><span data-ttu-id="e5817-102">OuvrirRequête, action de macro</span><span class="sxs-lookup"><span data-stu-id="e5817-102">OpenQuery Macro Action</span></span>


<span data-ttu-id="e5817-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="e5817-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="e5817-p101">Faites appel à l'action **OuvrirRequête** pour ouvrir une requête Sélection ou Analyse croisée en mode Feuille de données, Création ou Aperçu avant impression. Cette action exécute une requête Action. Vous pouvez également sélectionner le mode de saisie des données voulu pour la requête.</span><span class="sxs-lookup"><span data-stu-id="e5817-p101">You can use the **OpenQuery** action to open a select or crosstab query in Datasheet view, Design view, or Print Preview. This action runs an action query. You can also select a data entry mode for the query.</span></span>


> [!NOTE]
> <P><span data-ttu-id="e5817-p102">[!REMARQUE] Cette action n'est disponible que dans l'environnement de base de données Access (.mdb ou .accdb). Reportez-vous aux actions <STRONG>OuvrirVue</STRONG>, <STRONG>OuvrirProcédureStockée</STRONG> ou <STRONG>OuvrirFonction</STRONG> si vous utilisez l'environnement de projet Access (.adp).</span><span class="sxs-lookup"><span data-stu-id="e5817-p102">This action is only available in the Access database environment (.mdb or .accdb). See the <STRONG>OpenView</STRONG>, <STRONG>OpenStoredProcedure</STRONG>, or <STRONG>OpenFunction</STRONG> actions if you are using the Access project environment (.adp).</span></span></P>



## <a name="setting"></a><span data-ttu-id="e5817-109">Paramètre</span><span class="sxs-lookup"><span data-stu-id="e5817-109">Setting</span></span>

<span data-ttu-id="e5817-110">L'action **OuvrirRequête** possède les arguments suivants.</span><span class="sxs-lookup"><span data-stu-id="e5817-110">The **OpenQuery** action has the following arguments.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="e5817-111">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="e5817-111">Action argument</span></span></p></th>
<th><p><span data-ttu-id="e5817-112">Description</span><span class="sxs-lookup"><span data-stu-id="e5817-112">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="e5817-113"><strong>Nom de la requête</strong></span><span class="sxs-lookup"><span data-stu-id="e5817-113"><strong>Query Name</strong></span></span></p></td>
<td><p><span data-ttu-id="e5817-114">Nom de la requête à ouvrir.</span><span class="sxs-lookup"><span data-stu-id="e5817-114">The name of the query to open.</span></span> <span data-ttu-id="e5817-115">La zone <strong>Nom de la requête</strong> , dans la section <strong>Arguments de l’Action</strong> du volet Générateur de Macro affiche toutes les requêtes dans la base de données en cours.</span><span class="sxs-lookup"><span data-stu-id="e5817-115">The <strong>Query Name</strong> box in the <strong>Action Arguments</strong> section of the Macro Builder pane shows all queries in the current database.</span></span> <span data-ttu-id="e5817-116">Il s’agit d’un argument obligatoire. Si vous exécutez une macro contenant l’action <strong>OuvrirRequête</strong> dans une base de données bibliothèque, Microsoft Access recherche d’abord la requête portant ce nom dans la base de données bibliothèque, puis dans la base de données en cours.</span><span class="sxs-lookup"><span data-stu-id="e5817-116">This is a required argument.If you run a macro containing the <strong>OpenQuery</strong> action in a library database, Microsoft Access first looks for the query with this name in the library database, and then in the current database.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="e5817-117"><strong>View</strong></span><span class="sxs-lookup"><span data-stu-id="e5817-117"><strong>View</strong></span></span></p></td>
<td><p><span data-ttu-id="e5817-p104">Affichage dans lequel s’ouvre la requête. Cliquez sur <strong>Feuille de données</strong>, <strong>Création</strong>, <strong>Aperçu avant impression</strong>, <strong>Tableau croisé dynamique</strong> ou <strong>Graphique croisé dynamique</strong> dans la zone <strong>Affichage</strong>. La valeur par défaut est <strong>Feuille de données</strong>.</span><span class="sxs-lookup"><span data-stu-id="e5817-p104">The view in which the query will open. Click <strong>Datasheet</strong>, <strong>Design</strong>, <strong>Print Preview</strong>, <strong>PivotTable</strong>, or <strong>PivotChart</strong> in the <strong>View</strong> box. The default is <strong>Datasheet</strong>.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="e5817-121"><strong>Mode Données</strong></span><span class="sxs-lookup"><span data-stu-id="e5817-121"><strong>Data Mode</strong></span></span></p></td>
<td><p><span data-ttu-id="e5817-p105">Mode de saisie de données de la requête. S’applique uniquement aux requêtes ouvertes en mode Feuille de données. Cliquez sur <strong>Ajouter</strong> (l’utilisateur ne peut pas modifier les enregistrements existants, mais peut en ajouter de nouveaux), <strong>Modifier</strong> (l’utilisateur peut modifier les enregistrements existants et en ajouter de nouveaux) ou <strong>Lecture seule</strong> l’utilisateur peut uniquement consulter les enregistrements). La valeur par défaut est <strong>Modifier</strong>.</span><span class="sxs-lookup"><span data-stu-id="e5817-p105">The data entry mode for the query. This applies only to queries opened in Datasheet view. Click <strong>Add</strong> (the user can add new records but can't edit existing records), <strong>Edit</strong> (the user can edit existing records and add new records), or <strong>Read Only</strong> (the user can only view records). The default is <strong>Edit</strong>.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="e5817-126">Notes</span><span class="sxs-lookup"><span data-stu-id="e5817-126">Remarks</span></span>

<span data-ttu-id="e5817-127">Si l'argument **Affichage** a la valeur **Feuille de données**, Access affiche le jeu de résultats si la requête est une requête Sélection, Analyse croisée, Union ou SQL direct dont la propriété **RenvoieSur** a la valeur **Oui**; et renvoie la requête s'il s'agit d'une requête Action, Définition des données ou SQL direct dont la propriété **RenvoieSur** a la valeur **Non**.</span><span class="sxs-lookup"><span data-stu-id="e5817-127">If you use **Datasheet** for the **View** argument, Access displays the result set if the query is a select, crosstab, union, or pass-through query whose **ReturnsRecords** property is set to **Yes**; and it runs the query if it is an action, data-definition, or pass-through query whose **ReturnsRecords** property is set to **No**.</span></span>

<span data-ttu-id="e5817-p106">L'action **OuvrirRequête** équivaut à double-cliquer sur la requête dans le volet de navigation ou à cliquer avec le bouton droit sur la requête dans le volet de navigation et à choisir un affichage. Cette action vous permet de sélectionner des options supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="e5817-p106">The **OpenQuery** action is similar to double-clicking the query in the Navigation Pane, or right-clicking the query in the Navigation Pane and selecting a view. With this action you can select additional options.</span></span>

<span data-ttu-id="e5817-130">**Conseils**</span><span class="sxs-lookup"><span data-stu-id="e5817-130">**Tips**</span></span>

  - <span data-ttu-id="e5817-p107">Vous pouvez faire glisser une requête depuis le volet de navigation vers une ligne d'action de macro. Ceci crée automatiquement une action **OuvrirRequête** qui ouvre la requête en mode Feuille de données. Basculer en mode Création lorsque la requête est ouverte supprime le paramètre de l'argument **Mode Données** de la requête. Ce paramètre n'est pas actif, même si l'utilisateur revient en mode Feuille de données.</span><span class="sxs-lookup"><span data-stu-id="e5817-p107">You can drag a query from the Navigation Pane to a macro action row. This automatically creates an **OpenQuery** action that opens the query in Datasheet view. Switching to Design view while the query is open removes the **Data Mode** argument setting for the query. This setting isn't in effect even if the user returns to Datasheet view.</span></span>

  - <span data-ttu-id="e5817-135">Si vous ne voulez pas afficher les messages système qui s'affichent normalement lorsqu'une requête d'action est exécutée (indiquant qu'il s'agit d'une requête d'action et affichant le nombre d'enregistrements concernés), vous pouvez faire appel à l'action **Avertissements** pour supprimer l'affichage de ces messages.</span><span class="sxs-lookup"><span data-stu-id="e5817-135">If you don't want to display the system messages that normally appear when an action query is run (indicating it is an action query and showing how many records will be affected), you can use the **SetWarnings** action to suppress the display of these messages.</span></span>

<span data-ttu-id="e5817-136">Pour exécuter l'action **OuvrirRequête** dans un module Visual Basic pour Applications (VBA), utilisez la méthode **OpenQuery** de l'objet **DoCmd**.</span><span class="sxs-lookup"><span data-stu-id="e5817-136">To run the **OpenQuery** action in a Visual Basic for Applications (VBA) module, use the **OpenQuery** method of the **DoCmd** object.</span></span>
