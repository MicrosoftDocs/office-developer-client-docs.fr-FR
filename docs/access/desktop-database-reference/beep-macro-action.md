---
title: Bip, Action de Macro (référence de base de données du bureau Access)
TOCTitle: Beep Macro Action
ms:assetid: 5ca1600f-7934-3b3d-19fd-f305cda0e5d8
ms:mtpsurl: https://msdn.microsoft.com/library/Ff194572(v=office.15)
ms:contentKeyID: 48545092
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm11853
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 7024734b09544042fa58b8cd95127d8195e5788a
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471883"
---
# <a name="beep-macro-action"></a><span data-ttu-id="c630f-102">Bip, action de macro</span><span class="sxs-lookup"><span data-stu-id="c630f-102">Beep Macro Action</span></span>


<span data-ttu-id="c630f-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="c630f-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="c630f-104">Utilisez l'action **Bip** pour émettre un signal sonore dans les haut-parleurs de l'ordinateur.</span><span class="sxs-lookup"><span data-stu-id="c630f-104">You can use the **Beep** action to sound a beep tone through the computer's speaker.</span></span>

## <a name="setting"></a><span data-ttu-id="c630f-105">Valeur</span><span class="sxs-lookup"><span data-stu-id="c630f-105">Setting</span></span>

<span data-ttu-id="c630f-106">L'action **Bip** ne possède aucun argument.</span><span class="sxs-lookup"><span data-stu-id="c630f-106">The **Beep** action doesn't have any arguments.</span></span>

## <a name="remarks"></a><span data-ttu-id="c630f-107">Remarques</span><span class="sxs-lookup"><span data-stu-id="c630f-107">Remarks</span></span>

<span data-ttu-id="c630f-108">Vous pouvez utiliser l'action **Bip** pour signaler les occurrences suivantes :</span><span class="sxs-lookup"><span data-stu-id="c630f-108">You can use the **Beep** action to signal the following occurrences:</span></span>

  - <span data-ttu-id="c630f-109">des modifications importantes ont été apportées à l'écran ;</span><span class="sxs-lookup"><span data-stu-id="c630f-109">Important screen changes have occurred.</span></span>

  - <span data-ttu-id="c630f-p101">le type de données entré dans un contrôle est incorrect. Par exemple, l'utilisateur a entré des donnes numériques dans un contrôle zone de texte ;</span><span class="sxs-lookup"><span data-stu-id="c630f-p101">The wrong kind of data has been entered in a control. For example, the user has entered numeric data in a text box control.</span></span>

  - <span data-ttu-id="c630f-112">une macro a atteint un point spécifique ou a terminé ses actions.</span><span class="sxs-lookup"><span data-stu-id="c630f-112">A macro has reached a specified point or has completed its actions.</span></span>

<span data-ttu-id="c630f-113">La fréquence et la durée du signal sonore dépendent du matériel, qui peut varier d'un ordinateur à un autre.</span><span class="sxs-lookup"><span data-stu-id="c630f-113">The frequency and duration of the beep depend on the hardware, which may vary between computers.</span></span>

<span data-ttu-id="c630f-114">Pour exécuter l'action **Bip** dans un module Visual Basic pour Applications (VBA), utilisez la méthode **Beep** de l'objet **DoCmd**.</span><span class="sxs-lookup"><span data-stu-id="c630f-114">To run the **Beep** action in a Visual Basic for Applications (VBA) module, use the **Beep** method of the **DoCmd** object.</span></span>
