---
title: AfficherPointeurSablier, action de macro
TOCTitle: DisplayHourglassPointer Macro Action
ms:assetid: 2c93039a-f75c-abeb-1dfa-e632a5bdf6f2
ms:mtpsurl: https://msdn.microsoft.com/library/Ff192103(v=office.15)
ms:contentKeyID: 48543957
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm117200
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 003fb36fc876aa573419b963a9eca6f54332190a
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470766"
---
# <a name="displayhourglasspointer-macro-action"></a><span data-ttu-id="bfeba-102">AfficherPointeurSablier, action de macro</span><span class="sxs-lookup"><span data-stu-id="bfeba-102">DisplayHourglassPointer Macro Action</span></span>


<span data-ttu-id="bfeba-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="bfeba-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="bfeba-p101">Utilisez l'action **AfficherPointeurSablier** pour transformer le pointeur de la souris en icône de sablier (ou toute autre icône choisie) pendant l'exécution d'une macro. Cette action peut fournir une indication visuelle de l'exécution de la macro. Elle est particulièrement utile lorsque l'exécution d'une action de macro ou de la macro elle-même prend beaucoup de temps.</span><span class="sxs-lookup"><span data-stu-id="bfeba-p101">You can use the **DisplayHourglassPointer** action to change the mouse pointer to an image of an hourglass (or another icon you've chosen) while a macro is running. This action can provide a visual indication that the macro is running. This is especially useful when a macro action or the macro itself takes a long time to run.</span></span>

## <a name="setting"></a><span data-ttu-id="bfeba-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="bfeba-107">Setting</span></span>

<span data-ttu-id="bfeba-108">L'action **AfficherPointeurSablier** utilise l'argument suivant :</span><span class="sxs-lookup"><span data-stu-id="bfeba-108">The **DisplayHourglassPointer** action has the following argument.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="bfeba-109">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="bfeba-109">Action argument</span></span></p></th>
<th><p><span data-ttu-id="bfeba-110">Description</span><span class="sxs-lookup"><span data-stu-id="bfeba-110">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="bfeba-111"><strong>Sablier actif</strong></span><span class="sxs-lookup"><span data-stu-id="bfeba-111"><strong>Hourglass On</strong></span></span></p></td>
<td><p><span data-ttu-id="bfeba-p102">Cliquez sur <strong>Oui</strong> (afficher l’icône) ou sur <strong>Non</strong> (afficher le pointeur de souris normal) dans la zone <strong>Sablier actif</strong> de la section <strong>Arguments de l’action</strong> du volet Générateur de macro. La valeur par défaut est <strong>Oui</strong>.</span><span class="sxs-lookup"><span data-stu-id="bfeba-p102">Click <strong>Yes</strong> (display the icon) or <strong>No</strong> (display the normal mouse pointer) in the <strong>Hourglass On</strong> box in the <strong>Action Arguments</strong> section of the Macro Builder pane. The default is <strong>Yes</strong>.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="bfeba-114">Remarques</span><span class="sxs-lookup"><span data-stu-id="bfeba-114">Remarks</span></span>

<span data-ttu-id="bfeba-115">Cette action est souvent utilisée lorsque l'écho est désactivé à l'aide de l'action **Écho**.</span><span class="sxs-lookup"><span data-stu-id="bfeba-115">You often use this action if you have turned echo off by using the **Echo** action.</span></span> <span data-ttu-id="bfeba-116">Lorsque l’écho est désactivé, Access suspend les mises à jour de l’écran jusqu'à ce que la macro est terminée.</span><span class="sxs-lookup"><span data-stu-id="bfeba-116">When echo is off, Access suspends screen updates until the macro is finished.</span></span>

<span data-ttu-id="bfeba-117">Access réinitialise automatiquement l'argument **Sablier actif** sur la valeur **Non** une fois la macro terminée.</span><span class="sxs-lookup"><span data-stu-id="bfeba-117">Access automatically resets the **Hourglass On** argument to **No** when the macro finishes running.</span></span>


> [!NOTE]
> <UL>
> <LI>
> <P><span data-ttu-id="bfeba-p104">Dans Microsoft Windows, il s'agit de l'icône définie pour <STRONG>Occupé</STRONG> dans la boîte de dialogue <STRONG>Propriétés de Souris</STRONG> du Panneau de configuration de Windows. L'icône utilisée par défaut pour tous les systèmes d'exploitation Windows est un sablier animé.</span><span class="sxs-lookup"><span data-stu-id="bfeba-p104">In Microsoft Windows, this is the icon you set for <STRONG>Busy</STRONG> in the <STRONG>Mouse Properties</STRONG> dialog box of Windows Control Panel. The default for all Windows operating systems is an animated hourglass icon.</span></span></P>
> <LI>
> <P><span data-ttu-id="bfeba-120">Vous pouvez choisir toute autre icône.</span><span class="sxs-lookup"><span data-stu-id="bfeba-120">You can choose another icon if you want.</span></span></P></LI></UL>



<span data-ttu-id="bfeba-121">Pour exécuter l'action **AfficherPointeurSablier** dans un module Visual Basic pour Applications (VBA), utilisez la méthode **Hourglass** de l'objet **DoCmd**.</span><span class="sxs-lookup"><span data-stu-id="bfeba-121">To run the **DisplayHourglassPointer** action in a Visual Basic for Applications (VBA) module, use the **Hourglass** method of the **DoCmd** object.</span></span>
