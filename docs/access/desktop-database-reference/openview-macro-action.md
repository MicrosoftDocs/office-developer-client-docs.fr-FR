---
title: OuvrirVue, action de macro
TOCTitle: OpenView Macro Action
ms:assetid: 4d3b7e6d-4b81-4fbe-7222-24d745350321
ms:mtpsurl: https://msdn.microsoft.com/library/Ff193569(v=office.15)
ms:contentKeyID: 48544726
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm50135
f1_categories:
- Office.Version=v15
ms.openlocfilehash: ad607f852f5af21bc2979bd635286b86d18489a5
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470345"
---
# <a name="openview-macro-action"></a><span data-ttu-id="76140-102">OuvrirVue, action de macro</span><span class="sxs-lookup"><span data-stu-id="76140-102">OpenView Macro Action</span></span>


<span data-ttu-id="76140-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="76140-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="76140-p101">Dans un projet Access, vous pouvez utiliser l'action **OuvrirVue** pour ouvrir une vue en mode Feuille de données, Création ou Aperçu avant impression. Cette action exécute la vue nommée lorsqu'elle est ouverte en mode Feuille de données. Vous pouvez choisir le mode de saisie des données voulu pour la vue et limiter les enregistrements que celle-ci affiche.</span><span class="sxs-lookup"><span data-stu-id="76140-p101">In an Access project, you can use the **OpenView** action to open a view in Datasheet view, Design view, or Print Preview. This action runs the named view when opened in Datasheet view. You can select data entry for the view and restrict the records that the view displays.</span></span>


> [!NOTE]
> <P><span data-ttu-id="76140-p102">[!REMARQUE] Cette action ne sera pas autorisée si la base de données n'est pas approuvée. Pour plus d'informations sur l'activation des macros, voir les liens dans la section See Alsode cet article.</span><span class="sxs-lookup"><span data-stu-id="76140-p102">This action will not be allowed if the database is not trusted. For more information about enabling macros, see the links in the See Also section of this article.</span></span></P>



## <a name="setting"></a><span data-ttu-id="76140-109">Paramètre</span><span class="sxs-lookup"><span data-stu-id="76140-109">Setting</span></span>

<span data-ttu-id="76140-110">L'action **OuvrirVue** possède les arguments suivants.</span><span class="sxs-lookup"><span data-stu-id="76140-110">The **OpenView** action has the following arguments.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="76140-111">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="76140-111">Action argument</span></span></p></th>
<th><p><span data-ttu-id="76140-112">Description</span><span class="sxs-lookup"><span data-stu-id="76140-112">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="76140-113"><strong>Nom de la vue</strong></span><span class="sxs-lookup"><span data-stu-id="76140-113"><strong>View Name</strong></span></span></p></td>
<td><p><span data-ttu-id="76140-114">Nom de la vue à ouvrir.</span><span class="sxs-lookup"><span data-stu-id="76140-114">The name of the view to open.</span></span> <span data-ttu-id="76140-115">La zone <strong>Nom de l’affichage</strong> dans la section <strong>Arguments de l’Action</strong> du volet Générateur de Macro affiche toutes les vues dans la base de données en cours.</span><span class="sxs-lookup"><span data-stu-id="76140-115">The <strong>View Name</strong> box in the <strong>Action Arguments</strong> section of the Macro Builder pane shows all views in the current database.</span></span> <span data-ttu-id="76140-116">Cet argument est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="76140-116">This is a required argument.</span></span> <span data-ttu-id="76140-117">Si vous exécutez une macro contenant l’action <strong>OuvrirVue</strong> dans une base de données bibliothèque, Microsoft Access recherche d’abord la vue portant ce nom dans la base de données bibliothèque, puis dans la base de données active.</span><span class="sxs-lookup"><span data-stu-id="76140-117">If you run a macro containing the <strong>OpenView</strong> action in a library database, Microsoft Access first looks for the view with this name in the library database, and then in the current database.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="76140-118"><strong>View</strong></span><span class="sxs-lookup"><span data-stu-id="76140-118"><strong>View</strong></span></span></p></td>
<td><p><span data-ttu-id="76140-p104">Affichage dans lequel s’ouvre la vue. Cliquez sur <strong>Feuille de données</strong>, <strong>Création</strong>, <strong>Aperçu avant impression</strong>, <strong>Tableau croisé dynamique</strong> ou <strong>Graphique croisé dynamique</strong> dans la zone <strong>Affichage</strong>. La valeur par défaut est <strong>Feuille de données</strong>.</span><span class="sxs-lookup"><span data-stu-id="76140-p104">The view in which the view will open. Click <strong>Datasheet</strong>, <strong>Design</strong>, <strong>Print Preview</strong>, <strong>PivotTable</strong>, or <strong>PivotChart</strong> in the <strong>View</strong> box. The default is <strong>Datasheet</strong>.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="76140-122"><strong>Mode Données</strong></span><span class="sxs-lookup"><span data-stu-id="76140-122"><strong>Data Mode</strong></span></span></p></td>
<td><p><span data-ttu-id="76140-p105">Mode de saisie de données de la vue. S’applique uniquement aux vues ouvertes en mode Feuille de données. Cliquez sur <strong>Ajouter</strong> (l’utilisateur ne peut ni afficher ni modifier les enregistrements existants, mais peut en ajouter de nouveaux), <strong>Modifier</strong> (l’utilisateur peut afficher et modifier les enregistrements existants et en ajouter de nouveaux) ou <strong>Lecture seule</strong> (l’utilisateur peut uniquement consulter les enregistrements). La valeur par défaut est <strong>Modifier</strong>.</span><span class="sxs-lookup"><span data-stu-id="76140-p105">The data entry mode for the view. This applies only to views opened in Datasheet view. Click <strong>Add</strong> (the user can add new records but can't view or edit existing records), <strong>Edit</strong> (the user can view or edit existing records and add new records), or <strong>Read Only</strong> (the user can only view records). The default is <strong>Edit</strong>.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="76140-127">Notes</span><span class="sxs-lookup"><span data-stu-id="76140-127">Remarks</span></span>

<span data-ttu-id="76140-128">Cette action équivaut à double-cliquer sur une vue dans le volet de navigation ou à cliquer avec le bouton droit sur la vue dans le volet de navigation, puis à sélectionner la commande de votre choix.</span><span class="sxs-lookup"><span data-stu-id="76140-128">This action is similar to double-clicking a view in the Navigation Pane, or right-clicking the view in the Navigation Pane and then clicking the command you want.</span></span>

<span data-ttu-id="76140-129">**Conseils**</span><span class="sxs-lookup"><span data-stu-id="76140-129">**Tips**</span></span>

  - <span data-ttu-id="76140-p106">Vous pouvez faire glisser une vue depuis le volet de navigation vers une ligne d'action de macro. Ceci crée automatiquement une action **OuvrirVue** qui ouvre la vue en mode Feuille de données.</span><span class="sxs-lookup"><span data-stu-id="76140-p106">You can drag a view from the Navigation Pane to a macro action row. This automatically creates an **OpenView** action that opens the view in Datasheet view.</span></span>

  - <span data-ttu-id="76140-132">Si vous ne voulez pas afficher les messages système qui s'affichent normalement lorsqu'une vue est exécutée (indiquant qu'il s'agit d'une vue et affichant le nombre d'enregistrements concernés), vous pouvez faire appel à l'action **Avertissements** pour supprimer l'affichage de ces messages.</span><span class="sxs-lookup"><span data-stu-id="76140-132">If you don't want to display the system messages that normally appear when a view is run (indicating it is a view and showing how many records will be affected), you can use the **SetWarning** action to suppress the display of these messages.</span></span>

<span data-ttu-id="76140-133">Pour exécuter l'action **OuvrirVue** dans un module Visual Basic pour Applications (VBA), utilisez la méthode **OpenView** de l'objet **DoCmd**.</span><span class="sxs-lookup"><span data-stu-id="76140-133">To run the **OpenView** action in a Visual Basic for Applications (VBA) module, use the **OpenView** method of the **DoCmd** object.</span></span>
