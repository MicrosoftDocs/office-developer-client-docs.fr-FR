---
title: Masquer le ruban au démarrage d’Access
TOCTitle: Hide the ribbon when Access starts
ms:assetid: f98bab58-8094-1c56-f70b-ced2e7849574
ms:mtpsurl: https://msdn.microsoft.com/library/Ff837012(v=office.15)
ms:contentKeyID: 48548817
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 2b2b3e28157662a585fa03353da92a598b96bd2c
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469754"
---
# <a name="hide-the-ribbon-when-access-starts"></a><span data-ttu-id="439db-102">Masquer le ruban au démarrage d’Access</span><span class="sxs-lookup"><span data-stu-id="439db-102">Hide the ribbon when Access starts</span></span>

<span data-ttu-id="439db-103">**S’applique à :** Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="439db-103">**Applies to:** Access 2013 | Office 2013</span></span>

<span data-ttu-id="439db-p101">Par défaut, Microsoft Access ne fournit pas de méthode pour masquer le ruban. Cette rubrique décrit comment charger un ruban personnalisé qui masque tous les onglets prédéfinis.</span><span class="sxs-lookup"><span data-stu-id="439db-p101">By default, Microsoft Access does not provide a method for hiding the ribbon. This topic describes how to load a customized ribbon that hides all of the built-in tabs.</span></span>

<span data-ttu-id="439db-106">Pour charger le ruban personnalisé au démarrage d'Access, vous devez stocker ses paramètres dans une table dénommée **RubansSysU**.</span><span class="sxs-lookup"><span data-stu-id="439db-106">To load the customized ribbon when Access starts, you should store its settings in a table named **USysRibbons**.</span></span>

<span data-ttu-id="439db-p102">Cette **table** doit être créée en utilisant des noms de colonnes spécifiques afin que les personnalisations du ruban soient implémentées. La table ci-dessous contient les paramètres à utiliser lors de la création de la table **RubansSysU**.</span><span class="sxs-lookup"><span data-stu-id="439db-p102">The **USysRibbons** table must be created using specific column names in order for the ribbon customizations to be implemented. The following table lists the settings to use when creating the **USysRibbons** table.</span></span>

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="439db-109">Nom de colonne</span><span class="sxs-lookup"><span data-stu-id="439db-109">Column Name</span></span></p></th>
<th><p><span data-ttu-id="439db-110">Type de données</span><span class="sxs-lookup"><span data-stu-id="439db-110">Data Type</span></span></p></th>
<th><p><span data-ttu-id="439db-111">Description</span><span class="sxs-lookup"><span data-stu-id="439db-111">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="439db-112"><strong>RibbonName</strong></span><span class="sxs-lookup"><span data-stu-id="439db-112"><strong>RibbonName</strong></span></span></p></td>
<td><p><span data-ttu-id="439db-113">Texte</span><span class="sxs-lookup"><span data-stu-id="439db-113">Text</span></span></p></td>
<td><p><span data-ttu-id="439db-114">Contient le nom du ruban personnalisé devant être associé à cette personnalisation.</span><span class="sxs-lookup"><span data-stu-id="439db-114">Contains the name of the custom ribbon to be associated with this customization.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="439db-115"><strong>RibbonXML</strong></span><span class="sxs-lookup"><span data-stu-id="439db-115"><strong>RibbonXML</strong></span></span></p></td>
<td><p><span data-ttu-id="439db-116">Mémo</span><span class="sxs-lookup"><span data-stu-id="439db-116">Memo</span></span></p></td>
<td><p><span data-ttu-id="439db-117">Contient le XML d'extensibilité du ruban (RibbonX) qui définit la personnalisation du ruban.</span><span class="sxs-lookup"><span data-stu-id="439db-117">Contains the Ribbon Extensibility XML (RibbonX) that defines the ribbon customization.</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="439db-118">La table contient les paramètres de personnalisation du ruban à stocker dans la table **RubansSysU**.</span><span class="sxs-lookup"><span data-stu-id="439db-118">The following table lists the ribbon customization settings to store in the **USysRibbons** table.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="439db-119">Nom de colonne</span><span class="sxs-lookup"><span data-stu-id="439db-119">Column Name</span></span></p></th>
<th><p><span data-ttu-id="439db-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="439db-120">Value</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="439db-121"><strong>RibbonName</strong></span><span class="sxs-lookup"><span data-stu-id="439db-121"><strong>RibbonName</strong></span></span></p></td>
<td><p><span data-ttu-id="439db-122">HideTheRibbon</span><span class="sxs-lookup"><span data-stu-id="439db-122">HideTheRibbon</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="439db-123"><strong>RibbonXML</strong></span><span class="sxs-lookup"><span data-stu-id="439db-123"><strong>RibbonXML</strong></span></span></p></td>
<td><p><span data-ttu-id="439db-124">&lt;CustomUI xmlns =&quot;https://schemas.microsoft.com/office/2006/01/CustomUI&quot; &gt; &lt;ruban startFromScratch =&quot;true&quot;/&gt;&lt;/CustomUI&gt;</span><span class="sxs-lookup"><span data-stu-id="439db-124">&lt;CustomUI xmlns=&quot;https://schemas.microsoft.com/office/2006/01/CustomUI&quot;&gt; &lt;ribbon startFromScratch=&quot;true&quot;/&gt;&lt;/CustomUI&gt;</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="applying-a-custom-ribbon-when-access-starts"></a><span data-ttu-id="439db-125">Appliquer le ruban personnalisé au démarrage d'Access</span><span class="sxs-lookup"><span data-stu-id="439db-125">Applying a Custom Ribbon When Access Starts</span></span>

<span data-ttu-id="439db-126">Pour appliquer un ruban personnalisé disponible au démarrage de l'application, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="439db-126">To apply a custom ribbon so that it is available when the application starts, use the following procedure:</span></span>

1.  <span data-ttu-id="439db-127">Suivez le processus décrit précédemment pour rendre le ruban personnalisé disponible pour cette application.</span><span class="sxs-lookup"><span data-stu-id="439db-127">Follow the process described previously to make the customized ribbon available to the application.</span></span>

2.  <span data-ttu-id="439db-128">Redémarrez l'application.</span><span class="sxs-lookup"><span data-stu-id="439db-128">Close and then restart the application.</span></span>

3.  <span data-ttu-id="439db-129">Cliquez sur le **Bouton Microsoft Office**![O12FileMenuButton\_ZA10077102](media/access-file-menu-button.gif "O12FileMenuButton_ZA10077102") , puis sur **Options Access**.</span><span class="sxs-lookup"><span data-stu-id="439db-129">Click the **Microsoft Office Button**![O12FileMenuButton\_ZA10077102](media/access-file-menu-button.gif "O12FileMenuButton_ZA10077102") and then click **Access Options**.</span></span>

4.  <span data-ttu-id="439db-130">Cliquez sur l'option **Base de données active**, dans la section **Options de la barre d'outils et du ruban**, puis sur la liste **Nom du ruban** et sélectionnez **HideTheRibbon**.</span><span class="sxs-lookup"><span data-stu-id="439db-130">Click the **Current Database** option and then, in the **Ribbon and Toolbar Options** section, click the **Ribbon Name** list and select **HideTheRibbon**.</span></span>

5.  <span data-ttu-id="439db-131">Fermez er redémarrez l'application.</span><span class="sxs-lookup"><span data-stu-id="439db-131">Close and then restart the application.</span></span>

> [!NOTE]
> <span data-ttu-id="439db-132">[!REMARQUE] Pour plus d'informations sur l'interface du ruban dans d'autres applications Office, consultez l'article [Vue d'ensemble du ruban Office Fluent](https://docs.microsoft.com/office/vba/Library-Reference/Concepts/overview-of-the-office-fluent-ribbon).</span><span class="sxs-lookup"><span data-stu-id="439db-132">For more information about the ribbon UI in other Office applications, see [Overview of the Office Fluent Ribbon](https://docs.microsoft.com/office/vba/Library-Reference/Concepts/overview-of-the-office-fluent-ribbon).</span></span>

