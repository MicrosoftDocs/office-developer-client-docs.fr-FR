---
title: OuvrirEtat, action de macro
TOCTitle: OpenReport Macro Action
ms:assetid: cd35faf2-190d-ac48-cf59-81c1599eb764
ms:mtpsurl: https://msdn.microsoft.com/library/Ff834462(v=office.15)
ms:contentKeyID: 48547758
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm188079
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 06828474aee961dd6df02f29c7d046805a1c764c
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470223"
---
# <a name="openreport-macro-action"></a><span data-ttu-id="c3186-102">OuvrirEtat, action de macro</span><span class="sxs-lookup"><span data-stu-id="c3186-102">OpenReport Macro Action</span></span>

<span data-ttu-id="c3186-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="c3186-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="c3186-p101">Faites appel à l'action **OuvrirEtat** pour ouvrir un état en mode Création ou Aperçu avant impression ou pour envoyer l'état directement à l'imprimante. Vous pouvez également limiter les enregistrements qui sont imprimés dans l'état.</span><span class="sxs-lookup"><span data-stu-id="c3186-p101">You can use the **OpenReport** action to open a report in Design view or Print Preview, or to send the report directly to the printer. You can also restrict the records that are printed in the report.</span></span>

## <a name="setting"></a><span data-ttu-id="c3186-106">Paramètre</span><span class="sxs-lookup"><span data-stu-id="c3186-106">Setting</span></span>

<span data-ttu-id="c3186-107">L'action **OuvrirEtat** possède les arguments suivants.</span><span class="sxs-lookup"><span data-stu-id="c3186-107">The **OpenReport** action has the following arguments.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="c3186-108">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="c3186-108">Action argument</span></span></p></th>
<th><p><span data-ttu-id="c3186-109">Description</span><span class="sxs-lookup"><span data-stu-id="c3186-109">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="c3186-110">Nom de l’état</span><span class="sxs-lookup"><span data-stu-id="c3186-110">Report Name</span></span></p></td>
<td><p><span data-ttu-id="c3186-p102">Nom de l'état à ouvrir. La zone <strong>Nom de l'état</strong> de la section <strong>Arguments de l'action</strong> du volet <strong>Générateur de macro</strong> présente tous les états dans la base de données actuelle. Il s'agit d'un argument obligatoire. Si vous exécutez une macro contenant l'action OpenReport dans une base de données bibliothèque, Microsoft Access recherche d'abord l'état portant ce nom dans la base de données bibliothèque, puis dans la base de données actuelle.  </span><span class="sxs-lookup"><span data-stu-id="c3186-p102">The name of the report to open. The <strong>Report Name</strong> box in the <strong>Action Arguments</strong> section of the <strong>Macro Builder</strong> pane shows all reports in the current database. This is a required argument. If you run a macro containing the OpenReport action in a library database, Microsoft Access first looks for the report with this name in the library database, and then in the current database.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c3186-115">Vue</span><span class="sxs-lookup"><span data-stu-id="c3186-115">View</span></span></p></td>
<td><p><span data-ttu-id="c3186-p103">Affichage dans lequel s'ouvre l'état. Cliquez sur <strong>Imprimer</strong> (imprime l'état immédiatement), <strong>Création</strong> ou <strong>Aperçu avant impression</strong> dans la zone <strong>Affichage</strong>. La valeur par défaut est <strong>Imprimer</strong>.  </span><span class="sxs-lookup"><span data-stu-id="c3186-p103">The view in which the report will open. Click <strong>Print</strong> (print the report immediately), <strong>Design</strong>, or <strong>Print Preview</strong> in the <strong>View</strong> box. The default is <strong>Print</strong>.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c3186-119">Filter Name</span><span class="sxs-lookup"><span data-stu-id="c3186-119">Filter Name</span></span></p></td>
<td><p><span data-ttu-id="c3186-p104">Filtre qui limite les enregistrements de l'état. Vous pouvez entrer le nom d'une requête existante quelconque ou d'un filtre enregistré en tant que requête. La requête doit toutefois inclure tous les champs de l'état que vous ouvrez, ou sa propriété <strong>TousLesChamps</strong> doit avoir la valeur <strong>Oui</strong>.  </span><span class="sxs-lookup"><span data-stu-id="c3186-p104">A filter that restricts the report's records. You can enter the name of either an existing query or a filter that was saved as a query. However, the query must include all the fields in the report you are opening or have its <strong>OutputAllFields</strong> property set to <strong>Yes</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="c3186-123">Where Condition</span><span class="sxs-lookup"><span data-stu-id="c3186-123">Where Condition</span></span></p></td>
<td><p><span data-ttu-id="c3186-124">Clause ou expression WHERE SQL valable (sans le mot WHERE) utilisée par Access pour sélectionner des enregistrements dans la table ou requête sous-jacente de l'état.</span><span class="sxs-lookup"><span data-stu-id="c3186-124">A valid SQL WHERE clause (without the word WHERE) or expression that Access uses to select records from the report's underlying table or query.</span></span> <span data-ttu-id="c3186-125">Si vous sélectionnez un filtre avec l’argument nom du filtre, Access applique cette clause WHERE aux résultats du filtre.</span><span class="sxs-lookup"><span data-stu-id="c3186-125">If you select a filter with the Filter Name argument, Access applies this WHERE clause to the results of the filter.</span></span> <span data-ttu-id="c3186-126">Pour ouvrir un état et limiter ses enregistrements à ceux spécifiés par la valeur d’un contrôle d’un formulaire, utilisez l’expression suivante :</span><span class="sxs-lookup"><span data-stu-id="c3186-126">To open a report and restrict its records to those specified by the value of a control on a form, use the following expression:</span></span><br /><span data-ttu-id="c3186-127">
<strong>[</strong><em>nomchamp</em><strong>] = Forms![</strong><em>nomformulaire</em><strong>]![</strong><em>nomcontrôle formulaire</em><strong>]</strong></span><span class="sxs-lookup"><span data-stu-id="c3186-127">
<strong>[</strong><em>fieldname</em><strong>] = Forms![</strong><em>formname</em><strong>]![</strong><em>controlname on form</em><strong>]</strong></span></span><br />
<span data-ttu-id="c3186-128">Remplacez <em>fieldname</em> par le nom d’un champ dans la table ou requête sous-jacente du rapport que vous souhaitez ouvrir.</span><span class="sxs-lookup"><span data-stu-id="c3186-128">Replace <em>fieldname</em> with the name of a field in the underlying table or query of the report you want to open.</span></span> <span data-ttu-id="c3186-129">Remplacez <em>formname</em> et <em>NomContrôle formulaire</em> par le nom du formulaire et le contrôle du formulaire qui contient la valeur à laquelle les enregistrements du rapport doivent correspondre.</span><span class="sxs-lookup"><span data-stu-id="c3186-129">Replace <em>formname</em> and <em>controlname on form</em> with the name of the form and the control on the form that contains the value you want records in the report to match.</span></span></p>
<p><span data-ttu-id="c3186-130"><b>Remarque</b>: la longueur maximale de l’argument Condition Where est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="c3186-130"><b>NOTE</b>: The maximum length of the Where Condition argument is 255 characters.</span></span> <span data-ttu-id="c3186-131">Si vous devez entrer une clause SQL WHERE plus complexe et plus longue, utilisez la méthode <b>OpenReport</b> de l’objet <b>DoCmd</b> dans un Visual Basic pour le module d’Applications (VBA) à la place.</span><span class="sxs-lookup"><span data-stu-id="c3186-131">If you need to enter a more complex SQL WHERE clause longer than this, use the <b>OpenReport</b> method of the <b>DoCmd</b> object in a Visual Basic for Applications (VBA) module instead.</span></span> <span data-ttu-id="c3186-132">Vous pouvez entrer des instructions de clause SQL WHERE jusqu'à 32 768 caractères dans VBA.</span><span class="sxs-lookup"><span data-stu-id="c3186-132">You can enter SQL WHERE clause statements of up to 32,768 characters in VBA.</span></span></p>
</td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="c3186-133">Window Mode</span><span class="sxs-lookup"><span data-stu-id="c3186-133">Window Mode</span></span></p></td>
<td><p><span data-ttu-id="c3186-p108">Mode dans lequel s'ouvre l'état. Cliquez sur <strong>Standard</strong>, <strong>Masqué</strong>, <strong>Icône</strong> ou <strong>Boîte de dialogue</strong> dans la zone <strong>Mode Fenêtre</strong>. La valeur par défaut est <strong>Standard</strong>.  </span><span class="sxs-lookup"><span data-stu-id="c3186-p108">The mode in which the report will open. Click <strong>Normal</strong>, <strong>Hidden</strong>, <strong>Icon</strong>, or <strong>Dialog</strong> in the <strong>Window Mode</strong> box. The default is <strong>Normal</strong>.</span></span></p>
<p><span data-ttu-id="c3186-137"><b>Remarque</b>: paramètres de l’argument Mode fenêtre certains ne s’appliquent pas à l’aide des documents à onglets.</span><span class="sxs-lookup"><span data-stu-id="c3186-137"><b>NOTE</b>: Some Window Mode argument settings do not apply when using tabbed documents.</span></span> <span data-ttu-id="c3186-138">Pour activer des fenêtres superposées :</span><span class="sxs-lookup"><span data-stu-id="c3186-138">To switch to overlapping windows:</span></span>
<ol>
<li><p><span data-ttu-id="c3186-139">Cliquez sur <strong>Options</strong>.</span><span class="sxs-lookup"><span data-stu-id="c3186-139">Click <strong>Options</strong>.</span></span></p></li>
<li><p><span data-ttu-id="c3186-140">Dans la boîte dialogue <strong>Options Access</strong>, cliquez sur <strong>Base de données active</strong>.</span><span class="sxs-lookup"><span data-stu-id="c3186-140">In the <strong>Access Options</strong> dialog box, click <strong>Current Database</strong>.</span></span></p></li>
<li><p><span data-ttu-id="c3186-141">Dans la section <strong>Options de l'application</strong>, sous <strong>Options de la fenêtre Document</strong>, cliquez sur <strong>Fenêtres superposées</strong>.</span><span class="sxs-lookup"><span data-stu-id="c3186-141">In the <strong>Application Options</strong> section, under <strong>Document Window Options</strong>, click <strong>Overlapping Windows</strong>.</span></span></p></li>
<li><p><span data-ttu-id="c3186-142">Cliquez sur <strong>OK</strong>, puis fermez et rouvrez la base de données.</span><span class="sxs-lookup"><span data-stu-id="c3186-142">Click <strong>OK</strong>, and then close and reopen the database.</span></span></p></li>
</ol>
</td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="c3186-143">Remarques</span><span class="sxs-lookup"><span data-stu-id="c3186-143">Remarks</span></span>

<span data-ttu-id="c3186-p110">Le paramètre **Print** de l'argument **View** imprime immédiatement l'état selon les paramètres d'imprimante actuels, sans ouvrir la boîte de dialogue **Imprimer**. Vous pouvez également utiliser l'action **OpenReport** pour ouvrir et configurer un état, puis utiliser l'action PrintOut pour l'imprimer. Par exemple, vous pouvez modifier l'état ou utiliser l'action **PrintOut** pour modifier les paramètres d'imprimante avant l'impression.</span><span class="sxs-lookup"><span data-stu-id="c3186-p110">The **Print** setting for the **View** argument prints the report immediately by using the current printer settings, without bringing up the **Print** dialog box. You can also use the **OpenReport** action to open and set up a report and then use the PrintOut action to print it. For example, you may want to modify the report or use the **PrintOut** action to change the printer settings before you print.</span></span>

<span data-ttu-id="c3186-147">Le filtre et la condition WHERE que vous appliquez deviennent le paramètre de la propriété **Filtre** de l'état.</span><span class="sxs-lookup"><span data-stu-id="c3186-147">The filter and WHERE condition you apply become the setting of the report's **Filter** property.</span></span>

<span data-ttu-id="c3186-148">L'action **OuvrirEtat** équivaut à double-cliquer sur l'état dans le volet de navigation ou à cliquer avec le bouton droit sur l'état dans le volet de navigation et à choisir un affichage ou la commande **Imprimer**.</span><span class="sxs-lookup"><span data-stu-id="c3186-148">The **OpenReport** action is similar to double-clicking the report in the Navigation Pane, or right-clicking the report in the Navigation Pane and selecting a view or the **Print** command.</span></span>

> [!TIP] 
> - <span data-ttu-id="c3186-149">Pour imprimer des états similaires pour différentes séries de données, utilisez un filtre ou une clause WHERE pour limiter les enregistrements imprimés dans l'état.</span><span class="sxs-lookup"><span data-stu-id="c3186-149">To print similar reports for different sets of data, use a filter or a WHERE clause to restrict the records printed in the report.</span></span> <span data-ttu-id="c3186-150">Ensuite, modifiez la macro pour appliquer un autre filtre ou modifier l’argument Condition Where.</span><span class="sxs-lookup"><span data-stu-id="c3186-150">Then edit the macro to apply a different filter or change the Where Condition argument.</span></span>
> 
> - <span data-ttu-id="c3186-p112">Vous pouvez faire glisser un état depuis le volet de navigation vers une ligne d'action de macro. Ceci crée automatiquement une action **OuvrirEtat** qui ouvre l'état en mode État.</span><span class="sxs-lookup"><span data-stu-id="c3186-p112">You can drag a report from the Navigation Pane to a macro action row. This automatically creates an **OpenReport** action that opens the report in Report view.</span></span>

## <a name="example"></a><span data-ttu-id="c3186-153">Exemple</span><span class="sxs-lookup"><span data-stu-id="c3186-153">Example</span></span>

<span data-ttu-id="c3186-154">L'exemple suivant montre comment utiliser l'action OpenReport pour transmettre un paramètre qui permet de filtrer un état lorsqu'il est ouvert.</span><span class="sxs-lookup"><span data-stu-id="c3186-154">The following example shows how to use the OpenReport action to pass a parameter that filters a report as it is opened.</span></span> <span data-ttu-id="c3186-155">Le rapport **rptChapters** affiche les enregistrements de l’auteur spécifié en transmettant l’élément sélectionné dans la zone de liste déroulante **cboAuthors** pour le paramètre SelectedAuthor.</span><span class="sxs-lookup"><span data-stu-id="c3186-155">The **rptChapters** report displays the records for the specified author by passing the item selected in the **cboAuthors** combo box to the SelectedAuthor parameter.</span></span>

<span data-ttu-id="c3186-156">**Exemple de code fourni par** la [référence du programmeur Microsoft Access 2010](https://www.amazon.com/Microsoft-Access-2010-Programmers-Reference/dp/8126528125).</span><span class="sxs-lookup"><span data-stu-id="c3186-156">**Sample code provided by** the [Microsoft Access 2010 Programmer’s Reference](https://www.amazon.com/Microsoft-Access-2010-Programmers-Reference/dp/8126528125).</span></span>

```vb
    OpenReport
        Report Name rptChapters
        View Report
        Filter Name
        Where Condition
        Window Mode Normal
    
    Parameters
        SelectedAuthor =[cboAuthor]
```