---
title: Parcourir, action de Macro
TOCTitle: BrowseTo Macro Action
ms:assetid: b25e1cc6-c4ed-abd6-0285-94fc7dae0bdf
ms:mtpsurl: https://msdn.microsoft.com/library/Ff822020(v=office.15)
ms:contentKeyID: 48547167
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- vbaac10.chm35083
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 09e2b8238555a0d1b718f04c69f733039c1d8477
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472497"
---
# <a name="browseto-macro-action"></a><span data-ttu-id="21acd-102">Parcourir, action de Macro</span><span class="sxs-lookup"><span data-stu-id="21acd-102">BrowseTo Macro Action</span></span>

<span data-ttu-id="21acd-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="21acd-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="21acd-p101">Vous pouvez utiliser l'action **Parcourir** pour naviguer parmi des objets en place. Vous pouvez également modifier l'objet source d'un contrôle de sous-formulaire en spécifiant l'argument Chemin d'accès au contrôle de sous-formulaire. Utilisez **Parcourir** pour naviguer de formulaire1 à formulaire2 sans ouvrir de nouvelle fenêtre.</span><span class="sxs-lookup"><span data-stu-id="21acd-p101">You can use the **BrowseTo** action to navigate between objects in place. You can also change the source object of a subform control by specifying the Path to Subform Control argument. Use **BrowseTo** to navigate from form1 to form2 without opening up a new window.</span></span>

## <a name="setting"></a><span data-ttu-id="21acd-107">Paramètre</span><span class="sxs-lookup"><span data-stu-id="21acd-107">Setting</span></span>

<span data-ttu-id="21acd-108">L'action **Parcourir** utilise l'argument suivant :</span><span class="sxs-lookup"><span data-stu-id="21acd-108">The **BrowseTo** action has the following argument.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="21acd-109">Argument de l’action</span><span class="sxs-lookup"><span data-stu-id="21acd-109">Action argument</span></span></p></th>
<th><p><span data-ttu-id="21acd-110">Description</span><span class="sxs-lookup"><span data-stu-id="21acd-110">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="21acd-111">Type d'objet</span><span class="sxs-lookup"><span data-stu-id="21acd-111">Object Type</span></span></p></td>
<td><p><span data-ttu-id="21acd-112">Type d'objet vers lequel naviguer.</span><span class="sxs-lookup"><span data-stu-id="21acd-112">The object type to which to browse.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="21acd-113">Nom de l’objet</span><span class="sxs-lookup"><span data-stu-id="21acd-113">Object Name</span></span></p></td>
<td><p><span data-ttu-id="21acd-114">Objet qui se charge à l’intérieur du contrôle de sous-formulaire référencé par l’argument Chemin d’accès au contrôle de sous-formulaire.</span><span class="sxs-lookup"><span data-stu-id="21acd-114">The object that loads inside the subform control referenced by the Path to Subform Control argument.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="21acd-115">Chemin d’accès au contrôle de sous-formulaire</span><span class="sxs-lookup"><span data-stu-id="21acd-115">Path to Subform Control</span></span></p></td>
<td><p><span data-ttu-id="21acd-116">Si spécifié, le chemin d’accès à partir du formulaire principal de l’application du sous-formulaire cible contrôle qui charge l’objet spécifié par l’argument nom de l’objet.</span><span class="sxs-lookup"><span data-stu-id="21acd-116">If specified, the path from the main form of the application to the target subform control that loads the object specified by the Object Name argument.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="21acd-117">Condition Where</span><span class="sxs-lookup"><span data-stu-id="21acd-117">Where Condition</span></span></p></td>
<td><p><span data-ttu-id="21acd-118">Spécifié, remplace la condition Where de la source de l'enregistrement de l'objet.</span><span class="sxs-lookup"><span data-stu-id="21acd-118">If specified, replaces the Where condition of the object record source.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="21acd-119">Page</span><span class="sxs-lookup"><span data-stu-id="21acd-119">Page</span></span></p></td>
<td><p><span data-ttu-id="21acd-p102">Spécifié, définit la page du formulaire continu qui deviendra la page active. Cet argument est réservé au Web.</span><span class="sxs-lookup"><span data-stu-id="21acd-p102">If specified, sets the page of the continuous form that will be made the current page. This argument is Web only.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="21acd-122">Mode Données</span><span class="sxs-lookup"><span data-stu-id="21acd-122">Data Mode</span></span></p></td>
<td><p><span data-ttu-id="21acd-123">Spécifié, le mode d'entrée des données du formulaire.</span><span class="sxs-lookup"><span data-stu-id="21acd-123">If specified, the data entry mode of the form.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="21acd-124">Remarques</span><span class="sxs-lookup"><span data-stu-id="21acd-124">Remarks</span></span>

<span data-ttu-id="21acd-125">L'argument Chemin d'accès au contrôle de sous-formulaire doit être spécifié à l'aide de la syntaxe de l'exemple de code suivant :</span><span class="sxs-lookup"><span data-stu-id="21acd-125">The PathToSubFormControl argument must be specified using the syntax in the following code example:</span></span>

```vb
    Main Form.SubForm Ctrl 1>Form 2.SubForm Ctrl 2>Form 3.SubFormCtrl3
```

<span data-ttu-id="21acd-p103">Dans cet exemple, le Formulaire principal est le formulaire de niveau supérieur dans l'application cliente Access. L'argument Chemin d'accès au contrôle de sous-formulaire doit spécifier en alternance les noms des contrôles de sous-formulaire et formulaire menant du formulaire principal au contrôle de sous-formulaire qui est le conteneur de l'objet spécifié par l'argument Nom d'objet. Chaque contrôle de sous-formulaire spécifié doit être un contrôle sur le formulaire qui le précède. Le chemin d'accès doit se terminer par un contrôle de sous-formulaire.</span><span class="sxs-lookup"><span data-stu-id="21acd-p103">In this example, the Main Form is the top level form in the Access client application. The Path to Sub Form Control argument must alternately specify form and subform control names leading from the main form to the subform control that is the container of the object specified by the Object Name argument. Each subform control specified must be a control on the form that precedes it. The path must end with a subform control.</span></span>

## <a name="example"></a><span data-ttu-id="21acd-130">Exemple</span><span class="sxs-lookup"><span data-stu-id="21acd-130">Example</span></span>

<span data-ttu-id="21acd-131">L’exemple suivant montre comment utiliser l’action Parcourir pour ouvrir un état dans un contrôle de sous-formulaire ou dans un contrôle de navigation.</span><span class="sxs-lookup"><span data-stu-id="21acd-131">The following example shows how to use the BrowseTo action to open a report in a subform control or within a navigation control.</span></span>

<span data-ttu-id="21acd-132">**Exemple de code fourni par** la [référence du programmeur Microsoft Access 2010](https://www.amazon.com/Microsoft-Access-2010-Programmers-Reference/dp/8126528125).</span><span class="sxs-lookup"><span data-stu-id="21acd-132">**Sample code provided by** the [Microsoft Access 2010 Programmer’s Reference](https://www.amazon.com/Microsoft-Access-2010-Programmers-Reference/dp/8126528125).</span></span>

```vb
    OnError
        Go to Next
        Macro Name
    
    /* Try to load the report in the host form (frmAuthorsParameter)    */
    BrowseTo
        Object Type Report
        Object Name rptChapters
        Path to Subform Control frmAuthorsParameter.sfrmChild
        Where Condition
        Page
        Data Mode Edit
    
    Parameters
        SelectedAuthor =[cboAuthor]
    
    /* if this fails, try to load it in the navigation subform     */
    BrowseTo
        Object Type Report
        Object Name rptChapters
        Path to Subform Control frmMain.NavigationSubform>frmAuthorsParameter.sfrmChild
        Where Condition
        Page
        Data Mode Edit
    
    Parameters
        SelectedAuthor =[cboAuthor]
```


