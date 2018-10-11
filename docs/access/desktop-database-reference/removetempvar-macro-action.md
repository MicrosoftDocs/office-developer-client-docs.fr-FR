---
title: SupprimerVarTemp, action de macro
TOCTitle: RemoveTempVar Macro Action
ms:assetid: 7bcc5010-3e30-ecef-2c5d-a35e73c8e325
ms:mtpsurl: https://msdn.microsoft.com/library/Ff196352(v=office.15)
ms:contentKeyID: 48545822
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm147125
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 61d03f23e0fbe04cae49670366f76c02e81d694a
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471289"
---
# <a name="removetempvar-macro-action"></a><span data-ttu-id="efbcb-102">SupprimerVarTemp, action de macro</span><span class="sxs-lookup"><span data-stu-id="efbcb-102">RemoveTempVar Macro Action</span></span>


<span data-ttu-id="efbcb-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="efbcb-103">**Applies to**: Access 2013 | Office 2013</span></span>



<span data-ttu-id="efbcb-104">L'action **SupprimerVarTemp** permet de supprimer une seule variable temporaire que vous avez créée en utilisant l'action **DéfinirVarTemp**.</span><span class="sxs-lookup"><span data-stu-id="efbcb-104">You can use the **RemoveTempVar** action to remove a single temporary variable that you created by using the **SetTempVar** action.</span></span>

## <a name="setting"></a><span data-ttu-id="efbcb-105">Paramètre</span><span class="sxs-lookup"><span data-stu-id="efbcb-105">Setting</span></span>

<span data-ttu-id="efbcb-106">L'action **SupprimerVarTemp** possède l'argument suivant.</span><span class="sxs-lookup"><span data-stu-id="efbcb-106">The **RemoveTempVar** action has the following argument.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="efbcb-107">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="efbcb-107">Action argument</span></span></p></th>
<th><p><span data-ttu-id="efbcb-108">Description</span><span class="sxs-lookup"><span data-stu-id="efbcb-108">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="efbcb-109"><strong>Name</strong></span><span class="sxs-lookup"><span data-stu-id="efbcb-109"><strong>Name</strong></span></span></p></td>
<td><p><span data-ttu-id="efbcb-110">Entrez le nom de la variable temporaire que vous souhaitez supprimer.</span><span class="sxs-lookup"><span data-stu-id="efbcb-110">Enter the name of the temporary variable you want to remove.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="efbcb-111">Notes</span><span class="sxs-lookup"><span data-stu-id="efbcb-111">Remarks</span></span>

  - <span data-ttu-id="efbcb-p101">Vous pouvez définir jusqu'à 255 variables temporaires à la fois. Toute variable temporaire non supprimée reste en mémoire jusqu'à ce que la base de données soit fermée. Il est conseillé de supprimer les variables temporaires lorsque vous ne les utilisez plus.</span><span class="sxs-lookup"><span data-stu-id="efbcb-p101">You can have up to 255 temporary variables defined at one time. If you do not remove a temporary variable, it will remain in memory until you close the database. It is a good practice to remove temporary variables when you are finished using them.</span></span>

  - <span data-ttu-id="efbcb-115">Access supprime automatiquement toutes les variables temporaires lorsque vous fermez la base de données ou le projet.</span><span class="sxs-lookup"><span data-stu-id="efbcb-115">Access automatically removes all temporary variables when you close the database or project.</span></span>

  - <span data-ttu-id="efbcb-p102">Si le nom de la variable à supprimer est mal orthographié, Access n'affiche aucun message d'erreur. La variable concernée est conservée en mémoire tant que vous ne fermez pas la base de données.</span><span class="sxs-lookup"><span data-stu-id="efbcb-p102">If you misspell the name of the variable to be removed, Access does not display an error. The variable you wanted to remove will remain in memory until you close the database.</span></span>

  - <span data-ttu-id="efbcb-118">Si vous avez créé plusieurs variables temporaires et souhaitez toutes les supprimer en une fois, utilisez l'action **SupprimerToutesVarTemp**.</span><span class="sxs-lookup"><span data-stu-id="efbcb-118">If you have created more than one temporary variable and you want to remove them all at once, use the **RemoveAllTempVars** action.</span></span>

  - <span data-ttu-id="efbcb-119">Pour exécuter l'action **SupprimerVarTemp** dans un module VBA, utilisez la méthode **Remove** de l'objet **TempVars**.</span><span class="sxs-lookup"><span data-stu-id="efbcb-119">To run the **RemoveTempVar** action in a VBA module, use the **Remove** method of the **TempVars** object.</span></span>

## <a name="example"></a><span data-ttu-id="efbcb-120">Exemple</span><span class="sxs-lookup"><span data-stu-id="efbcb-120">Example</span></span>

<span data-ttu-id="efbcb-121">La macro suivante illustre comment créer une variable temporaire, l'utiliser dans une condition et une zone de message, puis la supprimer en utilisant l'action **SupprimerVarTemp**.</span><span class="sxs-lookup"><span data-stu-id="efbcb-121">The following macro demonstrates how to create a temporary variable, use it in a condition and a message box, and then remove the temporary variable by using the **RemoveTempVar** action.</span></span>

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="efbcb-122">Condition</span><span class="sxs-lookup"><span data-stu-id="efbcb-122">Condition</span></span></p></th>
<th><p><span data-ttu-id="efbcb-123">Action</span><span class="sxs-lookup"><span data-stu-id="efbcb-123">Action</span></span></p></th>
<th><p><span data-ttu-id="efbcb-124">Arguments</span><span class="sxs-lookup"><span data-stu-id="efbcb-124">Arguments</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p></p></td>
<td><p><span data-ttu-id="efbcb-125"><strong>DéfinirVarTemp</strong></span><span class="sxs-lookup"><span data-stu-id="efbcb-125"><strong>SetTempVar</strong></span></span></p></td>
<td><p><span data-ttu-id="efbcb-126"><strong>Nom</strong>: MaVar<strong>Expression</strong>: BEntrée (&quot;Entrez un nombre différent de zéro.&quot;)</span><span class="sxs-lookup"><span data-stu-id="efbcb-126"><strong>Name</strong>: MyVar<strong>Expression</strong>: InputBox(&quot;Enter a non-zero number.&quot;)</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="efbcb-127">[TempVars]![MaVar] &lt;&gt;0</span><span class="sxs-lookup"><span data-stu-id="efbcb-127">[TempVars]![MyVar]&lt;&gt;0</span></span></p></td>
<td><p><span data-ttu-id="efbcb-128"><strong>MessageBox</strong></span><span class="sxs-lookup"><span data-stu-id="efbcb-128"><strong>MessageBox</strong></span></span></p></td>
<td><p><span data-ttu-id="efbcb-129"><strong>Message</strong>: =&quot;vous avez entré &quot; &amp; [VarTemp] ! [MaVar] &amp; &quot;. &quot; <strong>Bip</strong>: <strong>YesType</strong>: <strong>informations</strong></span><span class="sxs-lookup"><span data-stu-id="efbcb-129"><strong>Message</strong>: =&quot;You entered &quot; &amp; [TempVars]![MyVar] &amp; &quot;.&quot;<strong>Beep</strong>: <strong>YesType</strong>: <strong>Information</strong></span></span></p></td>
</tr>
<tr class="odd">
<td><p></p></td>
<td><p><span data-ttu-id="efbcb-130"><strong>SupprimerVarTemp</strong></span><span class="sxs-lookup"><span data-stu-id="efbcb-130"><strong>RemoveTempVar</strong></span></span></p></td>
<td><p><span data-ttu-id="efbcb-131"><strong>Nom</strong>: MaVar</span><span class="sxs-lookup"><span data-stu-id="efbcb-131"><strong>Name</strong>: MyVar</span></span></p></td>
</tr>
</tbody>
</table>
