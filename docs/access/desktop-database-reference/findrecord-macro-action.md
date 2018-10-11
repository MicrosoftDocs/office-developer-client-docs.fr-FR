---
title: TrouverEnregistrement, action de macro
TOCTitle: FindRecord Macro Action
ms:assetid: 379e3dda-cb7d-a294-29b1-c6ce9a62ba8a
ms:mtpsurl: https://msdn.microsoft.com/library/Ff192494(v=office.15)
ms:contentKeyID: 48544199
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm7496
f1_categories:
- Office.Version=v15
ms.openlocfilehash: d62c80c18ffd091d71c88fc64c9bd5c60c580101
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472537"
---
# <a name="findrecord-macro-action"></a><span data-ttu-id="4a7b7-102">TrouverEnregistrement, action de macro</span><span class="sxs-lookup"><span data-stu-id="4a7b7-102">FindRecord Macro Action</span></span>


<span data-ttu-id="4a7b7-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="4a7b7-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="4a7b7-104">Vous pouvez utiliser l’action **TrouverEnregistrement** pour rechercher la première instance de données qui répondent aux critères spécifiés par les arguments **TrouverEnregistrement** .</span><span class="sxs-lookup"><span data-stu-id="4a7b7-104">You can use the **FindRecord** action to find the first instance of data that meets the criteria specified by the **FindRecord** arguments.</span></span> <span data-ttu-id="4a7b7-105">Ces données peuvent être dans l’enregistrement actif, dans un enregistrement suivant ou précédent ou dans le premier enregistrement.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-105">This data can be in the current record, in a succeeding or prior record, or in the first record.</span></span> <span data-ttu-id="4a7b7-106">Vous pouvez rechercher des enregistrements dans la table active feuille de données, requête, formulaire ou formulaire.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-106">You can find records in the active table datasheet, query datasheet, form datasheet, or form.</span></span>

## <a name="setting"></a><span data-ttu-id="4a7b7-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="4a7b7-107">Setting</span></span>

<span data-ttu-id="4a7b7-108">L'action **TrouverEnregistrement** possède les arguments suivants.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-108">The **FindRecord** action has the following arguments.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="4a7b7-109">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="4a7b7-109">Action argument</span></span></p></th>
<th><p><span data-ttu-id="4a7b7-110">Description</span><span class="sxs-lookup"><span data-stu-id="4a7b7-110">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="4a7b7-111"><strong>Rechercher</strong></span><span class="sxs-lookup"><span data-stu-id="4a7b7-111"><strong>Find What</strong></span></span></p></td>
<td><p><span data-ttu-id="4a7b7-112">Spécifie les données que vous souhaitez rechercher dans l’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-112">Specifies the data you want to find in the record.</span></span> <span data-ttu-id="4a7b7-113">Entrez le texte, nombre ou date que vous souhaitez rechercher ou tapez une expression précédée d’un signe égal (<strong>=</strong>), dans la zone <strong>Rechercher</strong> , dans la section <strong>Arguments de l’Action</strong> du volet Générateur de Macro.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-113">Enter the text, number, or date you want to find or type an expression, which is preceded by an equal sign (<strong>=</strong>), in the <strong>Find What</strong> box in the <strong>Action Arguments</strong> section of the Macro Builder pane.</span></span> <span data-ttu-id="4a7b7-114">Vous pouvez utiliser des caractères génériques.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-114">You can use wildcard characters.</span></span> <span data-ttu-id="4a7b7-115">Cet argument est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-115">This is a required argument.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="4a7b7-116"><strong>Match</strong></span><span class="sxs-lookup"><span data-stu-id="4a7b7-116"><strong>Match</strong></span></span></p></td>
<td><p><span data-ttu-id="4a7b7-p103">Spécifie où se trouvent les données dans le champ. Vous pouvez rechercher des données dans toute partie du champ (<strong>N’importe où dans le champ</strong>), des données qui remplissent tout le champ (<strong>Champ entier</strong>) ou des données situées au début du champ (<strong>Début de champ</strong>). La valeur par défaut est <strong>Champ entier</strong>.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p103">Specifies where the data is located in the field. You can specify a search for data in any part of the field (<strong>Any Part of Field</strong>), for data that fills the entire field (<strong>Whole Field</strong>), or for data located at the beginning of the field (<strong>Start of Field</strong>). The default is <strong>Whole Field</strong>.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="4a7b7-120"><strong>Respecter la casse</strong></span><span class="sxs-lookup"><span data-stu-id="4a7b7-120"><strong>Match Case</strong></span></span></p></td>
<td><p><span data-ttu-id="4a7b7-p104">Indique si la recherche respecte la casse. Cliquez sur <strong>Oui</strong> (effectuer une recherche qui respecte la casse) ou sur <strong>Non</strong> (effectuer la recherche sans respecter les majuscules et les minuscules). La valeur par défaut est <strong>Non</strong>.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p104">Specifies whether the search is case-sensitive. Click <strong>Yes</strong> (conduct a case-sensitive search) or <strong>No</strong> (search without matching uppercase and lowercase letters exactly). The default is <strong>No</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="4a7b7-124"><strong>Rechercher</strong></span><span class="sxs-lookup"><span data-stu-id="4a7b7-124"><strong>Search</strong></span></span></p></td>
<td><p><span data-ttu-id="4a7b7-p105">Spécifie si la recherche part de l’enregistrement actif et remonte jusqu’au début des enregistrements (<strong>Haut</strong>), descend jusqu’à la fin des enregistrements (<strong>Bas</strong>) ou descend jusqu’à la fin des enregistrements, puis repart du début des enregistrements jusqu’à l’enregistrement actif afin que tous les enregistrements soient inclus dans la recherche (<strong>Tous</strong>). La valeur par défaut est <strong>Tous</strong>.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p105">Specifies whether the search proceeds from the current record up to the beginning of the records (<strong>Up</strong>); down to the end of the records (<strong>Down</strong>); or down to the end of the records and then from the beginning of the records to the current record, so all records are searched (<strong>All</strong>). The default is <strong>All</strong>.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="4a7b7-127"><strong>Avec mise en forme</strong></span><span class="sxs-lookup"><span data-stu-id="4a7b7-127"><strong>Search As Formatted</strong></span></span></p></td>
<td><p><span data-ttu-id="4a7b7-128">Spécifie si la recherche inclut les données mises en forme.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-128">Specifies whether the search includes formatted data.</span></span> <span data-ttu-id="4a7b7-129">Cliquez sur <strong>Oui</strong> (Microsoft Office Access 2007 recherche les données qu’il est mis en forme et affiché dans le champ) ou sur <strong>non</strong> (Access recherche les données qu’il est stocké dans la base de données, ce qui n’est pas toujours les mêmes tel qu’il est affiché).</span><span class="sxs-lookup"><span data-stu-id="4a7b7-129">Click <strong>Yes</strong> (Microsoft Office Access 2007 searches for the data as it is formatted and displayed in the field) or <strong>No</strong> (Access searches for the data as it is stored in the database, which isn't always the same as it's displayed).</span></span> <span data-ttu-id="4a7b7-130">La valeur par défaut est <strong>Non</strong>.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-130">The default is <strong>No</strong>.</span></span> <span data-ttu-id="4a7b7-131">Vous pouvez utiliser cette fonctionnalité pour limiter la recherche à des données dans un format particulier.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-131">You can use this feature to restrict the search to data in a particular format.</span></span> <span data-ttu-id="4a7b7-132">Par exemple, cliquez sur <strong>Oui</strong> , tapez <strong>1 234</strong> dans l’argument <strong>Rechercher</strong> pour trouver la valeur 1 234 dans un champ formaté afin d’inclure des virgules.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-132">For example, click <strong>Yes</strong> and type <strong>1,234</strong> in the <strong>Find What</strong> argument to find a value of 1,234 in a field formatted to include commas.</span></span> <span data-ttu-id="4a7b7-133">Cliquez sur <strong>non</strong> si vous souhaitez entrer <strong>1234</strong> pour rechercher des données dans ce champ.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-133">Click <strong>No</strong> if you want to type <strong>1234</strong> to search for the data in this field.</span></span> <span data-ttu-id="4a7b7-134">Pour rechercher des dates, cliquez sur <strong>Oui</strong> pour rechercher une date exactement comme elle est formatée, comme 08-juillet-2003.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-134">To search for dates, click <strong>Yes</strong> to find a date exactly as it is formatted, such as 08-July-2003.</span></span> <span data-ttu-id="4a7b7-135">Si vous cliquez sur <strong>non</strong>, entrez la date pour l’argument <strong>Rechercher</strong> dans le format défini dans les paramètres régionaux du Panneau de configuration Windows.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-135">If you click <strong>No</strong>, enter the date for the <strong>Find What</strong> argument in the format that is set in the regional settings in Windows Control Panel.</span></span> <span data-ttu-id="4a7b7-136">Ce format est indiqué dans la zone <strong>format de date court</strong> de l’onglet <strong>Date</strong> dans les paramètres régionaux.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-136">This format is shown in the <strong>Short date format</strong> box found on the <strong>Date</strong> tab in the regional settings.</span></span> <span data-ttu-id="4a7b7-137">Par exemple, si la zone <strong>format de date court</strong> est définie sur <strong>jj/aa</strong>, vous pouvez entrer 7/8/03, et Access trouve toutes les entrées dans un champ Date correspondant au 8 juillet 2003, quelle que soit la façon dont ce champ est mis en forme.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-137">For example, if the <strong>Short date format</strong> box is set to <strong>M/d/yy</strong>, you can enter 7/8/03, and Access will find all entries in a Date field that correspond to July 8, 2003, regardless of how this field is formatted.</span></span></p>

> [!NOTE]
> <P><span data-ttu-id="4a7b7-138">L’argument <STRONG>Avec mise en forme</STRONG> est appliqué uniquement si le champ actif est un contrôle dépendant, si l’argument <STRONG>Où</STRONG> est défini sur <STRONG>Champ entier</STRONG>, si l’argument <STRONG>Champ actif uniquement</STRONG> est défini sur <STRONG>Oui</STRONG> et si l’argument <STRONG>Respecter la casse</STRONG> est défini sur <STRONG>Non</STRONG>.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-138">The <STRONG>Search As Formatted</STRONG> argument takes effect only if the current field is a bound control, the <STRONG>Match</STRONG> argument is set to <STRONG>Whole Field</STRONG>, the <STRONG>Only Current Field</STRONG> argument is set to <STRONG>Yes</STRONG>, and the <STRONG>Match Case</STRONG> argument is set to <STRONG>No</STRONG>.</span></span></P>


<p><span data-ttu-id="4a7b7-139">Si vous définissez <strong>Respecter la casse</strong> sur <strong>Oui</strong> ou <strong>Champ actif uniquement</strong> sur <strong>Non</strong>, vous devez également définir <strong>Avec mise en forme</strong> sur <strong>Oui</strong>.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-139">If you set <strong>Match Case</strong> to <strong>Yes</strong> or <strong>Only Current Field</strong> to <strong>No</strong>, you must also set <strong>Search As Formatted</strong> to <strong>Yes</strong>.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="4a7b7-140"><strong>Champ actif uniquement</strong></span><span class="sxs-lookup"><span data-stu-id="4a7b7-140"><strong>Only Current Field</strong></span></span></p></td>
<td><p><span data-ttu-id="4a7b7-p107">Spécifie si la recherche est limitée au champ actif de chaque enregistrement ou si elle inclut tous les champs de chaque enregistrement. La recherche dans le champ actif est plus rapide. Cliquez sur <strong>Oui</strong> (limiter la recherche au champ actif) ou sur <strong>Non</strong> (rechercher dans tous les champs de chaque enregistrement). La valeur par défaut est <strong>Oui</strong>.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p107">Specifies whether the search is confined to the current field in each record or includes all fields in each record. Searching in the current field is faster. Click <strong>Yes</strong> (confine the search to the current field) or <strong>No</strong> (search in all fields in each record). The default is <strong>Yes</strong>.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="4a7b7-145"><strong>Trouver le premier</strong></span><span class="sxs-lookup"><span data-stu-id="4a7b7-145"><strong>Find First</strong></span></span></p></td>
<td><p><span data-ttu-id="4a7b7-p108">Indique si la recherche débute au premier enregistrement ou à l’enregistrement actif. Cliquez sur <strong>Oui</strong> (démarrer au premier enregistrement) ou sur <strong>Non</strong> (démarrer à l’enregistrement actif). La valeur par défaut est <strong>Oui</strong>.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p108">Specifies whether the search starts at the first record or at the current record. Click <strong>Yes</strong> (start at the first record) or <strong>No</strong> (start at the current record). The default is <strong>Yes</strong>.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="4a7b7-149">Remarques</span><span class="sxs-lookup"><span data-stu-id="4a7b7-149">Remarks</span></span>

<span data-ttu-id="4a7b7-p109">Lorsqu'une macro exécute l'action **TrouverEnregistrement**, Access recherche les données spécifiées dans les enregistrements (l'ordre de la recherche est déterminé par le paramètre de l'argument **Sens** ). Lorsqu'Access trouve les données spécifiées, elles sont sélectionnées dans l'enregistrement.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p109">When a macro runs the **FindRecord** action, Access searches for the specified data in the records (the order of the search is determined by the setting of the **Search** argument). When Access finds the specified data, the data is selected in the record.</span></span>

<span data-ttu-id="4a7b7-p110">L'action **TrouverEnregistrement** équivaut à cliquer sur **Rechercher** sous l'onglet **Accueil** et ses arguments sont identiques aux options de la boîte de dialogue **Rechercher et remplacer**. Si vous définissez les arguments de l'action **TrouverEnregistrement** dans le volet Générateur de macro, puis exécutez la macro, les options correspondantes sont sélectionnées dans la boîte de dialogue **Rechercher et remplacer** lorsque vous cliquez sur **Rechercher**.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p110">The **FindRecord** action is equivalent to clicking **Find** on the **Home** tab, and its arguments are the same as the options in the **Find and Replace** dialog box. If you set the **FindRecord** arguments in the Macro Builder pane and then run the macro, you will see the corresponding options selected in the **Find and Replace** dialog box when you click **Find**.</span></span>

<span data-ttu-id="4a7b7-p111">Access conserve les arguments les plus récents de l'action **TrouverEnregistrement** au cours d'une session de base de données afin que vous n'ayez pas à réentrer les mêmes critères si vous utilisez plusieurs fois l'action **TrouverEnregistrement**. Si vous laissez un argument vide, Access utilise le paramètre le plus récent de cet argument, qu'il ait été défini par une action **TrouverEnregistrement** précédente ou dans la boîte de dialogue **Rechercher et remplacer**.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p111">Access retains the most recent **FindRecord** arguments during a database session so that you don't need to enter the same criteria repeatedly as you perform subsequent operations with the **FindRecord** action. If you leave an argument blank, Access uses the most recent setting for the argument, as set either by a previous **FindRecord** action or in the **Find and Replace** dialog box.</span></span>

<span data-ttu-id="4a7b7-156">Si vous souhaitez rechercher un enregistrement à l'aide d'une macro, utilisez l'action **TrouverEnregistrement**, et non l'action **ExécuterCommandeMenu** avec ses arguments définis pour exécuter la commande **Rechercher**.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-156">When you want to find a record by using a macro, use the **FindRecord** action, not the **RunMenuCommand** action with its argument set to run the **Find** command.</span></span>


> [!NOTE]
> <P><span data-ttu-id="4a7b7-p112">[!REMARQUE] Bien que l'action <STRONG>TrouverEnregistrement</STRONG> corresponde à la commande <STRONG>Rechercher</STRONG> sous l'onglet <STRONG>Accueil</STRONG> pour les tables, les requêtes et les formulaires, elle ne correspond pas à la commande <STRONG>Rechercher</STRONG> dans le menu <STRONG>Edition</STRONG> de la fenêtre Code. Vous ne pouvez pas utiliser l'action <STRONG>TrouverEnregistrement</STRONG> pour rechercher du texte dans les modules.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p112">While the <STRONG>FindRecord</STRONG> action corresponds to the <STRONG>Find</STRONG> command on the <STRONG>Home</STRONG> tab for tables, queries, and forms, it doesn't correspond to the <STRONG>Find</STRONG> command on the <STRONG>Edit</STRONG> menu in the Code window. You can't use the <STRONG>FindRecord</STRONG> action to search for text in modules.</span></span></P>



<span data-ttu-id="4a7b7-p113">Si le texte actuellement sélectionné correspond au texte recherché au moment de l'exécution de l'action **TrouverEnregistrement**, la recherche commence immédiatement après la sélection, dans le même champ que la sélection et dans le même enregistrement. Sinon, la recherche commence au début de l'enregistrement actif. Cela vous permet de rechercher plusieurs instances des mêmes critères de recherche dans un même enregistrement.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p113">If the currently selected text is the same as the search text at the time the **FindRecord** action is carried out, the search begins immediately following the selection in the same field as the selection, and in the same record. Otherwise, the search begins at the start of the current record. This enables you to find multiple instances of the same search criteria that might appear in a single record.</span></span>

<span data-ttu-id="4a7b7-p114">Notez toutefois que si vous utilisez un bouton de commande pour exécuter une macro contenant l'action **TrouverEnregistrement**, la première instance des critères de recherche sera trouvée en boucle. En effet, le fait de cliquer sur le bouton de commande supprime le focus du champ contenant la valeur correspondante. L'action **TrouverEnregistrement** commencera la recherche au début de l'enregistrement. Pour éviter cela, exécutez la macro à l'aide d'une technique qui ne modifie pas le focus, comme un bouton de barre d'outils personnalisé ou une combinaison de touches définie dans une macro AutoKeys. Vous pouvez également définir le focus dans la macro sur le champ contenant les critères de recherche avant d'exécuter l'action **TrouverEnregistrement**.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p114">However, note that if you use a command button to run a macro containing the **FindRecord** action, the first instance of the search criteria will be found repeatedly. This behavior occurs because clicking the command button removes the focus from the field containing the matching value. The **FindRecord** action will then begin searching from the start of the record. To avoid this problem, run the macro by using a technique that doesn't change the focus, such as a custom toolbar button or a key combination defined in an AutoKeys macro, or set the focus in the macro to the field containing the search criteria before you carry out the **FindRecord** action.</span></span>

<table>
<thead>
<tr class="header">
<th><img src="media/access-alert-security.gif" title="Remarque sur la sécurité" alt="Security note" /><span data-ttu-id="4a7b7-167">de sécurité\*\*</span><span class="sxs-lookup"><span data-stu-id="4a7b7-167"><strong>Security Note</strong></span></span></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><span data-ttu-id="4a7b7-p115">
      Évitez d’utiliser l’instruction <strong>SendKeys</strong> ou une macro AutoKeys avec des informations sensibles et confidentielles. Un utilisateur malveillant pourrait intercepter la frappe et compromettre la sécurité de votre ordinateur et de vos données.
</span><span class="sxs-lookup"><span data-stu-id="4a7b7-p115">Avoid using the <strong>SendKeys</strong> statement or an AutoKeys macro with sensitive or confidential information. A malicious user could intercept the keystrokes and compromise the security of your computer and data.</span></span></td>
</tr>
</tbody>
</table>


<span data-ttu-id="4a7b7-170">Ce comportement est également observé si vous utilisez un bouton de commande pour exécuter une macro contenant l'action **TrouverSuivant**.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-170">The same behavior also occurs if you use a command button to run a macro containing the **FindNext** action.</span></span>

<span data-ttu-id="4a7b7-171">Pour exécuter l'action **TrouverEnregistrement** dans un module Visual Basic pour Applications (VBA), utilisez la méthode **FindRecord** de l'objet **DoCmd**.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-171">To run the **FindRecord** action in a Visual Basic for Applications (VBA) module, use the **FindRecord** method of the **DoCmd** object.</span></span>

<span data-ttu-id="4a7b7-172">Pour les recherches plus complexes, vous voudrez peut-être utiliser l'action de macro **RechercherEnregistrement**.</span><span class="sxs-lookup"><span data-stu-id="4a7b7-172">For more complex searches, you may want to use the **SearchForRecord** macro action.</span></span>
