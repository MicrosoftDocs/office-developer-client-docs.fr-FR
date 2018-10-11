---
title: SupprimerObjet, action de macro
TOCTitle: DeleteObject Macro Action
ms:assetid: a8deb2a7-4e73-8696-b8c1-3a3939d813f7
ms:mtpsurl: https://msdn.microsoft.com/library/Ff821415(v=office.15)
ms:contentKeyID: 48546912
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm152112
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 85e9fc888e06a69be6f458ed03ad92b8253b30a2
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470587"
---
# <a name="deleteobject-macro-action"></a><span data-ttu-id="40acf-102">SupprimerObjet, action de macro</span><span class="sxs-lookup"><span data-stu-id="40acf-102">DeleteObject Macro Action</span></span>


<span data-ttu-id="40acf-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="40acf-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="40acf-104">Utilisez l'action **SupprimerObjet** pour supprimer un objet de base de données spécifique.</span><span class="sxs-lookup"><span data-stu-id="40acf-104">You can use the **DeleteObject** action to delete a specified database object.</span></span>


> [!NOTE]
> <P><span data-ttu-id="40acf-p101">[!REMARQUE] Cette action ne sera pas autorisée si la base de données n'est pas approuvée. Pour plus d'informations sur l'activation des macros, voir les liens dans la section See Alsode cet article.</span><span class="sxs-lookup"><span data-stu-id="40acf-p101">This action will not be allowed if the database is not trusted. For more information about enabling macros, see the links in the See Also section of this article.</span></span></P>



## <a name="setting"></a><span data-ttu-id="40acf-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="40acf-107">Setting</span></span>

<span data-ttu-id="40acf-108">L'action **SupprimerObjet** possède les arguments suivants.</span><span class="sxs-lookup"><span data-stu-id="40acf-108">The **DeleteObject** action has the following arguments.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="40acf-109">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="40acf-109">Action argument</span></span></p></th>
<th><p><span data-ttu-id="40acf-110">Description</span><span class="sxs-lookup"><span data-stu-id="40acf-110">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="40acf-111"><strong>Type d'objet</strong></span><span class="sxs-lookup"><span data-stu-id="40acf-111"><strong>Object Type</strong></span></span></p></td>
<td><p><span data-ttu-id="40acf-p102">Type d’objet à supprimer. Cliquez sur <strong>Table</strong>, <strong>Requête</strong>, <strong>Formulaire</strong>, <strong>État</strong>, <strong>Macro</strong>, <strong>Module</strong>, <strong>Page d’accès aux données</strong>, <strong>Vue serveur</strong>, <strong>Schéma</strong>, <strong>Procédure stockée</strong> ou <strong>Fonction</strong> dans la zone <strong>Type d’objet</strong> de la section <strong>Arguments de l’action</strong> du volet Générateur de macro. Pour supprimer l’objet sélectionné dans le volet de navigation, laissez cet argument vide.</span><span class="sxs-lookup"><span data-stu-id="40acf-p102">The type of object to delete. Click <strong>Table</strong>, <strong>Query</strong>, <strong>Form</strong>, <strong>Report</strong>, <strong>Macro</strong>, <strong>Module</strong>, <strong>Data Access Page</strong>, <strong>Server View</strong>, <strong>Diagram</strong>, <strong>Stored Procedure</strong>, or <strong>Function</strong> in the <strong>Object Type</strong> box in the <strong>Action Arguments</strong> section of the Macro Builder pane. To delete the object selected in the Navigation Pane, leave this argument blank.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="40acf-115"><strong>Nom de l'objet</strong></span><span class="sxs-lookup"><span data-stu-id="40acf-115"><strong>Object Name</strong></span></span></p></td>
<td><p><span data-ttu-id="40acf-116">Nom de l’objet à supprimer.</span><span class="sxs-lookup"><span data-stu-id="40acf-116">The name of the object to delete.</span></span> <span data-ttu-id="40acf-117">La zone <strong>Nom de l'objet</strong> regroupe tous les objets de la base de données ayant le type sélectionné par l'argument <strong>Type d'objet</strong>.</span><span class="sxs-lookup"><span data-stu-id="40acf-117">The <strong>Object Name</strong> box shows all objects in the database of the type selected by the <strong>Object Type</strong> argument.</span></span> <span data-ttu-id="40acf-118">Si vous laissez la zone <strong>Type d’objet</strong> vide, laissez cette zone vide également.</span><span class="sxs-lookup"><span data-stu-id="40acf-118">If you leave the <strong>Object Type</strong> box blank, leave this box blank also.</span></span> <span data-ttu-id="40acf-119">Si vous exécutez une macro contenant l’action <strong>SupprimerObjet</strong> dans une base de données bibliothèque, Microsoft Office Access 2007 commence par rechercher un objet portant ce nom dans la base de données bibliothèque, puis dans la base de données active.</span><span class="sxs-lookup"><span data-stu-id="40acf-119">If you run a macro containing the <strong>DeleteObject</strong> action in a library database, Microsoft Office Access 2007 first looks for the object with this name in the library database, and then in the current database.</span></span></p></td>
</tr>
</tbody>
</table>



> [!WARNING]
> <P><span data-ttu-id="40acf-120">[!ATTENTION] Si vous laissez les zones <STRONG>Type d'objet</STRONG> et <STRONG>Nom de l'objet</STRONG> vides, Access supprime l'objet sélectionné dans le volet de navigation sans afficher de message d'avertissement lorsqu'il rencontre l'action <STRONG>SupprimerObjet</STRONG>.</span><span class="sxs-lookup"><span data-stu-id="40acf-120">If you leave the <STRONG>Object Type</STRONG> and <STRONG>Object Name</STRONG> boxes blank, Access deletes the object selected in the Navigation Pane without displaying a warning message when it encounters the <STRONG>DeleteObject</STRONG> action.</span></span></P>



## <a name="remarks"></a><span data-ttu-id="40acf-121">Remarques</span><span class="sxs-lookup"><span data-stu-id="40acf-121">Remarks</span></span>

<span data-ttu-id="40acf-p104">Vous pouvez utiliser l'action **SupprimerObjet** pour supprimer des objets temporaires créés pendant l'exécution de la macro. Par exemple, vous pouvez être amené à utiliser l'action **OuvrirRequête** pour exécuter une requête Création de table qui crée une table temporaire. Lorsque vous avez terminé d'utiliser la table temporaire, vous pouvez utiliser l'action **SupprimerObjet** pour la supprimer.</span><span class="sxs-lookup"><span data-stu-id="40acf-p104">You can use the **DeleteObject** action to delete temporary objects you have created while running the macro. For example, you could use the **OpenQuery** action to run a make-table query that creates a temporary table. When you are finished using the temporary table, you can use the **DeleteObject** action to delete it.</span></span>

<span data-ttu-id="40acf-125">Cette action équivaut à sélectionner un objet dans le volet de navigation, puis à appuyer sur la touche Suppr ou encore, à cliquer avec le bouton droit sur l'objet dans le volet de navigation, puis à cliquer sur **Supprimer**.</span><span class="sxs-lookup"><span data-stu-id="40acf-125">This action has the same effect as selecting an object in the Navigation Pane and then pressing the DEL key, or right-clicking the object in the Navigation Pane and clicking **Delete**.</span></span>

<span data-ttu-id="40acf-126">Pour exécuter l'action **SupprimerObjet** dans un module Visual Basic pour Applications, vous pouvez utiliser la méthode **DeleteObject** de l'objet **DoCmd**.</span><span class="sxs-lookup"><span data-stu-id="40acf-126">To run the **DeleteObject** action in a Visual Basic for Applications module, you can use the **DeleteObject** method of the **DoCmd** object.</span></span>
