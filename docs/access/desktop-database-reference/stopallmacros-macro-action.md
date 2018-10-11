---
title: ArrêtToutesMacros, action de macro
TOCTitle: StopAllMacros Macro Action
ms:assetid: 6afbf906-03b8-6e68-bbc9-7a4b141cf1c5
ms:mtpsurl: https://msdn.microsoft.com/library/Ff195440(v=office.15)
ms:contentKeyID: 48545442
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm104968
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 80da04f7b5f99fd0b249164caaeaca9f25edd172
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469156"
---
# <a name="stopallmacros-macro-action"></a><span data-ttu-id="5780f-102">ArrêtToutesMacros, action de macro</span><span class="sxs-lookup"><span data-stu-id="5780f-102">StopAllMacros Macro Action</span></span>


<span data-ttu-id="5780f-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="5780f-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="5780f-104">L'action **ArrêtToutesMacros** permet d'arrêter toutes les macros en cours d'exécution.</span><span class="sxs-lookup"><span data-stu-id="5780f-104">You can use the **StopAllMacros** action to stop all macros that are currently running.</span></span>

## <a name="setting"></a><span data-ttu-id="5780f-105">Valeur</span><span class="sxs-lookup"><span data-stu-id="5780f-105">Setting</span></span>

<span data-ttu-id="5780f-106">L'action **ArrêtToutesMacros** n'utilise aucun argument.</span><span class="sxs-lookup"><span data-stu-id="5780f-106">The **StopAllMacros** action doesn't have any arguments.</span></span>

## <a name="remarks"></a><span data-ttu-id="5780f-107">Remarques</span><span class="sxs-lookup"><span data-stu-id="5780f-107">Remarks</span></span>

<span data-ttu-id="5780f-108">Cette action est généralement utilisée lorsqu'une condition d'erreur nécessite l'arrêt de toutes les macros.</span><span class="sxs-lookup"><span data-stu-id="5780f-108">You typically use this action when an error condition makes it necessary to stop all macros.</span></span> <span data-ttu-id="5780f-109">Vous pouvez utiliser une expression conditionnelle dans la ligne d'action de la macro contenant cette action.</span><span class="sxs-lookup"><span data-stu-id="5780f-109">You can use a conditional expression in the macro's action row that contains this action.</span></span> <span data-ttu-id="5780f-110">Lorsque l’expression a la valeur **True** (– 1), Microsoft Access arrête toutes les macros.</span><span class="sxs-lookup"><span data-stu-id="5780f-110">When the expression evaluates to **True** (–1), Microsoft Access stops all macros.</span></span>

<span data-ttu-id="5780f-p102">C'est le cas, par exemple, lorsque l'affichage d'une boîte de message constitue l'une des actions complexes d'une macro, parmi lesquelles l'exécution d'autres macros. Si l'utilisateur clique sur **Annuler** dans cette boîte de message, l'action **ArrêtToutesMacros** peut entraîner l'arrêt de toutes les macros en cours d'exécution.</span><span class="sxs-lookup"><span data-stu-id="5780f-p102">For example, you might have a macro that displays a message box as one of a number of complex actions, including running other macros. If the user clicks **Cancel** in this message box, the **StopAllMacros** action can stop all the macros that are running.</span></span>

<span data-ttu-id="5780f-113">Si les actions **Écho** ou **Avertissements** ont été utilisées par une macro pour désactiver un écho ou l'affichage des messages système, l'action **ArrêtToutesMacros** les réactivera automatiquement.</span><span class="sxs-lookup"><span data-stu-id="5780f-113">If a macro has used the **Echo** or **SetWarnings** actions to turn echo or the display of system messages off, the **StopAllMacros** action automatically turns them back on.</span></span>

<span data-ttu-id="5780f-114">Cette action n'est pas disponible dans les modules Visual Basic pour Applications (VBA).</span><span class="sxs-lookup"><span data-stu-id="5780f-114">This action isn't available in a Visual Basic for Applications (VBA) module.</span></span>
