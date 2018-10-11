---
title: SurErreur, action de macro
TOCTitle: OnError Macro Action
ms:assetid: 5c6073c4-2c0f-0ed2-83b0-477636e2d81c
ms:mtpsurl: https://msdn.microsoft.com/library/Ff194562(v=office.15)
ms:contentKeyID: 48545088
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm62274
f1_categories:
- Office.Version=v15
ms.openlocfilehash: b27ffa119d37cd7ddf80292acd9a1e1e131b0359
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470561"
---
# <a name="onerror-macro-action"></a><span data-ttu-id="f2ac2-102">SurErreur, action de macro</span><span class="sxs-lookup"><span data-stu-id="f2ac2-102">OnError Macro Action</span></span>

<span data-ttu-id="f2ac2-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="f2ac2-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="f2ac2-104">Utilisez l'action **SurErreur** pour indiquer ce qui doit se passer lorsqu'une erreur se produit dans une macro.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-104">You can use the **OnError** action to specify what should happen when an error occurs in a macro.</span></span>

## <a name="setting"></a><span data-ttu-id="f2ac2-105">Valeur</span><span class="sxs-lookup"><span data-stu-id="f2ac2-105">Setting</span></span>

<span data-ttu-id="f2ac2-106">L'action **SurErreur** possède les arguments suivants.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-106">The **OnError** action has the following arguments.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="f2ac2-107">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="f2ac2-107">Action argument</span></span></p></th>
<th><p><span data-ttu-id="f2ac2-108">Description</span><span class="sxs-lookup"><span data-stu-id="f2ac2-108">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="f2ac2-109">Atteindre</span><span class="sxs-lookup"><span data-stu-id="f2ac2-109">Go to</span></span></p></td>
<td><p><span data-ttu-id="f2ac2-p101">Spécifie le comportement général lorsqu’une erreur se produit. Cliquez sur la flèche déroulante, puis sur l’un des paramètres suivants :</span><span class="sxs-lookup"><span data-stu-id="f2ac2-p101">Specify the general behavior that should occur when an error is encountered. Click the drop-down arrow and then click one of the following settings:</span></span></p>
<div class="tableSection">
<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="f2ac2-112">Valeur</span><span class="sxs-lookup"><span data-stu-id="f2ac2-112">Setting</span></span></p></th>
<th><p><span data-ttu-id="f2ac2-113">Description</span><span class="sxs-lookup"><span data-stu-id="f2ac2-113">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="f2ac2-114"><strong>Suivant</strong></span><span class="sxs-lookup"><span data-stu-id="f2ac2-114"><strong>Next</strong></span></span></p></td>
<td><p><span data-ttu-id="f2ac2-p102">Microsoft Office Access 2007 enregistre les détails de l’erreur dans l’objet <strong>MacroError</strong> mais n’arrête pas la macro. La macro passe à l’action suivante.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-p102">Microsoft Office Access 2007 records the details of the error in the <strong>MacroError</strong> object but does not stop the macro. The macro continues with the next action.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="f2ac2-117"><strong>Nom de la macro</strong></span><span class="sxs-lookup"><span data-stu-id="f2ac2-117"><strong>Macro Name</strong></span></span></p></td>
<td><p><span data-ttu-id="f2ac2-118">Access arrête l’exécution de la macro active et exécute la macro dont le nom figure dans l’argument <strong>Nom de la macro</strong>.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-118">Access stops the current macro and runs the macro that is named in the <strong>Macro Name</strong> argument.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="f2ac2-119"><strong>Échec</strong></span><span class="sxs-lookup"><span data-stu-id="f2ac2-119"><strong>Fail</strong></span></span></p></td>
<td><p><span data-ttu-id="f2ac2-120">Access arrête l'exécution de la macro active et affiche un message d'erreur.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-120">Access stops the current macro and displays an error message.</span></span></p></td>
</tr>
</tbody>
</table>

</div></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="f2ac2-121">Nom de la macro</span><span class="sxs-lookup"><span data-stu-id="f2ac2-121">Macro Name</span></span></p></td>
<td><p><span data-ttu-id="f2ac2-122">Si l’argument Atteindre est défini sur le nom de la Macro, tapez le nom de la macro à utiliser pour la gestion des erreurs.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-122">If the Go to argument is set to Macro Name, type the name of the macro to be used for error handling.</span></span> <span data-ttu-id="f2ac2-123">Le nom que vous tapez doit correspondre à un nom dans la colonne <strong>Nom de la Macro</strong> de la macro active. Vous ne pouvez pas entrer le nom d’un autre objet macro.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-123">The name you type must match a name in the <strong>Macro Name</strong> column of the current macro; you can't enter the name of a different macro object.</span></span> <span data-ttu-id="f2ac2-124">Dans l’exemple ci-dessous, la macro <strong>SurErreur</strong> est contenue dans le même objet de macro que l’action <strong>SurErreur</strong> .</span><span class="sxs-lookup"><span data-stu-id="f2ac2-124">In the example below, the <strong>ErrorHandler</strong> macro is contained in the same macro object as the <strong>OnError</strong> action.</span></span> <span data-ttu-id="f2ac2-125">Cet argument doit rester vide si l’argument Atteindre est défini sur <strong>suivant</strong> ou <strong>Échec</strong>.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-125">This argument must be left blank if the Go to argument is set to <strong>Next</strong> or <strong>Fail</strong>.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="f2ac2-126">Remarques</span><span class="sxs-lookup"><span data-stu-id="f2ac2-126">Remarks</span></span>

- <span data-ttu-id="f2ac2-p104">L'action **SurErreur** est généralement placée au début d'une macro, mais vous pouvez également la placer plus loin dans la macro. Les règles établies par cette action prendront effet dès exécution de l'action.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-p104">The **OnError** action is usually placed at the beginning of a macro, but you can also place the action later in the macro. The rules established by the action will take effect whenever the action is run.</span></span>

- <span data-ttu-id="f2ac2-p105">Si vous définissez l'argument Atteindre sur **Échec**, Access se comporte de la même manière que si la macro ne contenait pas d'action **SurErreur**. Autrement dit, si une erreur se produit, Access arrête la macro et affiche un message d'erreur standard. Le paramètre **Échec** est principalement utilisé pour désactiver la gestion des erreurs établie précédemment dans une macro.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-p105">If you set the Go to argument to **Fail**, Access behaves the same way it would if there were no **OnError** action in the macro. That is, if an error is encountered, Access stops the macro and displays a standard error message. The main use for the **Fail** setting is to turn off any error handling that you established earlier in a macro.</span></span>

## <a name="example"></a><span data-ttu-id="f2ac2-132">Exemple</span><span class="sxs-lookup"><span data-stu-id="f2ac2-132">Example</span></span>

<span data-ttu-id="f2ac2-133">La macro suivante illustre l'utilisation de l'action **SurErreur**.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-133">The following macro demonstrates the use of the **OnError** action.</span></span> <span data-ttu-id="f2ac2-134">Dans cet exemple, l'action **SurErreur** spécifie qu'Access exécute une macro de gestion des erreurs personnalisée nommée GestionnaireErreur lorsqu'une erreur se produit.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-134">In this example, the **OnError** action specifies that Access run a custom error handling macro named ErrorHandler when an error occurs.</span></span> <span data-ttu-id="f2ac2-135">Lorsqu’une erreur se produit, le submacro CatchErrors est appelée.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-135">When an error occurs, the CatchErrors submacro is called.</span></span> <span data-ttu-id="f2ac2-136">Si le numéro d’erreur est 2102, un message s’affiche et l’exécution de la macro est interrompue.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-136">If the error number is 2102, a specific message is displayed and macro execution is halted.</span></span> <span data-ttu-id="f2ac2-137">Sinon, un message décrivant l’erreur s’affiche et la macro est interrompue de sorte que vous pouvez effectuer des opérations de dépannage.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-137">Otherwise, a message describing the error is displayed and the macro is paused so that you can perform additional troubleshooting.</span></span> <span data-ttu-id="f2ac2-138">Cette macro affiche une zone de message qui fait référence à l'objet **MacroError** pour afficher des informations sur l'erreur.</span><span class="sxs-lookup"><span data-stu-id="f2ac2-138">The ErrorHandler macro displays a message box that refers to the **MacroError** object to display information about the error.</span></span>

<span data-ttu-id="f2ac2-139">**Exemple de code fourni par** la [référence du programmeur Microsoft Access 2010](https://www.amazon.com/Microsoft-Access-2010-Programmers-Reference/dp/8126528125).</span><span class="sxs-lookup"><span data-stu-id="f2ac2-139">**Sample code provided by** the [Microsoft Access 2010 Programmer’s Reference](https://www.amazon.com/Microsoft-Access-2010-Programmers-Reference/dp/8126528125).</span></span>

```vb
    /* MACRO: mcrThrowErrors                                  */
    /* PURPOSE: Error handling using macros in Access 2010    */
    
    OnError
        Go to Macro Name
        Macro Name CatchErrors
    
    OpenForm 
        Form Name frmSamples
        View Form
        Filter Name
        Where Condition
        Data Mode
        Window Mode Normal
    
    MessageBox 
        Message This message appears after the OpenForm action
        Beep Yes
        Type None
        Title
    
    
    /* SUBMACRO: CatchErrors                                   */
    
    SubMacro: CatchErrors
        If [MacroError].[Number]=2101 Then
            MessageBox
                Message Cannot find the specified form!
                Beep Yes
                Type Critical
                Title
            StopMacro
    
        Else
            MessageBox
                Message =[MacroErro].[Description]
                Beep Yes
                Type None
                Title Unhandled Error
    
            SingleStep
        End If
    
    End SubMacro
```