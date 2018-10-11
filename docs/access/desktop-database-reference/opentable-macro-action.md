---
title: OuvrirTable, action de macro
TOCTitle: OpenTable Macro Action
ms:assetid: 4220ad3a-d064-0034-2806-ec1a447cebac
ms:mtpsurl: https://msdn.microsoft.com/library/Ff192909(v=office.15)
ms:contentKeyID: 48544469
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm149011
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 729c1f14d49b3f80ec6307ce775066d3291b96de
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470944"
---
# <a name="opentable-macro-action"></a><span data-ttu-id="96c17-102">OuvrirTable, action de macro</span><span class="sxs-lookup"><span data-stu-id="96c17-102">OpenTable Macro Action</span></span>


<span data-ttu-id="96c17-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="96c17-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="96c17-p101">Faites appel à l'action **OuvrirTable** pour ouvrir une table en mode Feuille de données, Création ou Aperçu avant impression. Vous pouvez également sélectionner le mode de saisie des données voulu pour la table.</span><span class="sxs-lookup"><span data-stu-id="96c17-p101">You can use the **OpenTable** action to open a table in Datasheet view, Design view, or Print Preview. You can also select a data entry mode for the table.</span></span>

## <a name="setting"></a><span data-ttu-id="96c17-106">Paramètre</span><span class="sxs-lookup"><span data-stu-id="96c17-106">Setting</span></span>

<span data-ttu-id="96c17-107">L'action **OuvrirTable** possède les arguments suivants.</span><span class="sxs-lookup"><span data-stu-id="96c17-107">The **OpenTable** action has the following arguments.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="96c17-108">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="96c17-108">Action argument</span></span></p></th>
<th><p><span data-ttu-id="96c17-109">Description</span><span class="sxs-lookup"><span data-stu-id="96c17-109">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="96c17-110"><strong>Nom de la table</strong></span><span class="sxs-lookup"><span data-stu-id="96c17-110"><strong>Table Name</strong></span></span></p></td>
<td><p><span data-ttu-id="96c17-111">Le nom de la table à ouvrir.</span><span class="sxs-lookup"><span data-stu-id="96c17-111">The name of the table to open.</span></span> <span data-ttu-id="96c17-112">La zone <strong>Nom de la Table</strong> dans la section <strong>Arguments de l’Action</strong> du volet Générateur de Macro affiche toutes les tables dans la base de données en cours.</span><span class="sxs-lookup"><span data-stu-id="96c17-112">The <strong>Table Name</strong> box in the <strong>Action Arguments</strong> section of the Macro Builder pane shows all tables in the current database.</span></span> <span data-ttu-id="96c17-113">Cet argument est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="96c17-113">This is a required argument.</span></span> <span data-ttu-id="96c17-114">Si vous exécutez une macro contenant l’action <strong>OuvrirTable</strong> dans une base de données bibliothèque, Microsoft Access recherche d’abord la table portant ce nom dans la base de données bibliothèque, puis dans la base de données active.</span><span class="sxs-lookup"><span data-stu-id="96c17-114">If you run a macro containing the <strong>OpenTable</strong> action in a library database, Microsoft Microsoft Access looks for the table with this name first in the library database, then in the current database.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="96c17-115"><strong>View</strong></span><span class="sxs-lookup"><span data-stu-id="96c17-115"><strong>View</strong></span></span></p></td>
<td><p><span data-ttu-id="96c17-p103">Affichage dans lequel s’ouvre la table. Cliquez sur <strong>Feuille de données</strong>, <strong>Création</strong>, <strong>Aperçu avant impression</strong>, <strong>Tableau croisé dynamique</strong> ou <strong>Graphique croisé dynamique</strong> dans la zone <strong>Affichage</strong>. La valeur par défaut est <strong>Feuille de données</strong>.</span><span class="sxs-lookup"><span data-stu-id="96c17-p103">The view in which the table will open. Click <strong>Datasheet</strong>, <strong>Design</strong>, <strong>Print Preview</strong>, <strong>PivotTable</strong>, or <strong>PivotChart</strong> in the <strong>View</strong> box. The default is <strong>Datasheet</strong>.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="96c17-119"><strong>Mode Données</strong></span><span class="sxs-lookup"><span data-stu-id="96c17-119"><strong>Data Mode</strong></span></span></p></td>
<td><p><span data-ttu-id="96c17-p104">Mode de saisie de données de la table. S’applique uniquement aux tables ouvertes en mode Feuille de données. Cliquez sur <strong>Ajouter</strong> (l’utilisateur ne peut pas modifier les enregistrements existants, mais peut en ajouter de nouveaux), <strong>Modifier</strong> (l’utilisateur peut modifier les enregistrements existants et en ajouter de nouveaux) ou <strong>Lecture seule</strong> l’utilisateur peut uniquement consulter les enregistrements). La valeur par défaut est <strong>Modifier</strong>.</span><span class="sxs-lookup"><span data-stu-id="96c17-p104">The data entry mode for the table. This applies only to tables opened in Datasheet view. Click <strong>Add</strong> (the user can add new records but can't edit existing records), <strong>Edit</strong> (the user can edit existing records and add new records), or <strong>Read Only</strong> (the user can only view records). The default is <strong>Edit</strong>.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="96c17-124">Notes</span><span class="sxs-lookup"><span data-stu-id="96c17-124">Remarks</span></span>

<span data-ttu-id="96c17-125">Cette action équivaut à double-cliquer sur la table dans le volet de navigation ou à cliquer dessus avec le bouton droit dans le volet de navigation, puis à choisir un affichage.</span><span class="sxs-lookup"><span data-stu-id="96c17-125">This action is similar to double-clicking the table in the Navigation Pane, or right-clicking the table in the Navigation Pane and selecting a view.</span></span>


> [!TIP]
> <P><span data-ttu-id="96c17-p105">[!CONSEIL] Vous pouvez faire glisser une table depuis le volet de navigation vers une ligne d'action de macro. Ceci crée automatiquement une action <STRONG>OuvrirTable</STRONG> qui ouvre la table en mode Feuille de données.</span><span class="sxs-lookup"><span data-stu-id="96c17-p105">You can drag a table from the Navigation Pane to a macro action row. This automatically creates an <STRONG>OpenTable</STRONG> action that opens the table in Datasheet view.</span></span></P>



<span data-ttu-id="96c17-128">Pour exécuter l'action **OuvrirTable** dans un module Visual Basic pour Applications (VBA), utilisez la méthode **OpenTable** de l'objet **DoCmd**.</span><span class="sxs-lookup"><span data-stu-id="96c17-128">To run the **OpenTable** action in a Visual Basic for Applications (VBA) module, use the **OpenTable** method of the **DoCmd** object.</span></span>
