---
title: Détermination des éléments pris en charge
TOCTitle: Determining What is Supported
ms:assetid: 47b44dc9-e0fd-f204-0c68-e0de9247ee2c
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249221(v=office.15)
ms:contentKeyID: 48544602
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 159f1fae8e0d0c18f4c274b292f9cdc48c691b0e
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470362"
---
# <a name="determining-what-is-supported"></a><span data-ttu-id="cbba1-102">Détermination des éléments pris en charge</span><span class="sxs-lookup"><span data-stu-id="cbba1-102">Determining What is Supported</span></span>


<span data-ttu-id="cbba1-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="cbba1-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="cbba1-p101">La méthode **Supports** permet de déterminer si un objet **Recordset** spécifié prend en charge un type donné de fonctionnalité. Sa syntaxe est la suivante :</span><span class="sxs-lookup"><span data-stu-id="cbba1-p101">The **Supports** method is used to determine whether a specified **Recordset** object supports a particular type of functionality. It has the following syntax:</span></span>

`boolean = recordset.Supports( CursorOptions )`

<span data-ttu-id="cbba1-p102">La méthode **Supports** retourne une valeur booléenne indiquant si toutes les fonctionnalités identifiées par l'argument *CursorOptions* sont prises en charge par le fournisseur. Vous pouvez utiliser la méthode **Supports** pour déterminer les types de fonctionnalités pris en charge par un objet **Recordset**. Si l'objet **Recordset** prend en charge les fonctionnalités dont les constantes correspondantes se trouvent dans *CursorOptions*, la méthode **Supports** retourne la valeur **True**. Sinon, elle retourne la valeur **False**.</span><span class="sxs-lookup"><span data-stu-id="cbba1-p102">**Supports** returns a Boolean value that indicates whether all of the features identified by the *CursorOptions* argument are supported by the provider. You can use the **Supports** method to determine what types of functionality a **Recordset** object supports. If the **Recordset** object supports the features whose corresponding constants are in *CursorOptions*, the **Supports** method returns **True**. Otherwise, it returns **False**.</span></span>

<span data-ttu-id="cbba1-p103">À l'aide de la méthode **Supports**, vous pouvez vérifier si l'objet **Recordset** peut ajouter de nouveaux enregistrements, utiliser des signets, utiliser la méthode **Find**, utiliser le défilement, utiliser la propriété **Index** et procéder à des mises à jour par lot. Pour obtenir une liste complète des constantes et de leur signification, consultez [CursorOptionEnum](cursoroptionenum.md).</span><span class="sxs-lookup"><span data-stu-id="cbba1-p103">Using the **Supports** method, you can check for the ability of the **Recordset** object to add new records, use bookmarks, use the **Find** method, use scrolling, use the **Index** property, and to perform batch updates. For a complete list of constants and their meanings, see [CursorOptionEnum](cursoroptionenum.md).</span></span>

<span data-ttu-id="cbba1-p104">Même si la méthode **Supports** retourne la valeur **True** pour une fonctionnalité donnée, rien ne garantit que le fournisseur puisse la fournir en toutes circonstances. La valeur retournée par la méthode **Supports** indique simplement si le fournisseur peut prendre en charge la fonctionnalité spécifiée, pour autant que certaines conditions soient remplies. Par exemple, la méthode **Supports** peut indiquer qu'un objet **Recordset** prend en charge les mises à jour, même si le curseur dépend d'une jointure de plusieurs tables  dont certaines colonnes ne sont pas modifiables.</span><span class="sxs-lookup"><span data-stu-id="cbba1-p104">Although the **Supports** method may return **True** for a given functionality, it does not guarantee that the provider can make the feature available under all circumstances. The **Supports** method simply returns whether the provider can support the specified functionality, assuming certain conditions are met. For example, the **Supports** method may indicate that a **Recordset** object supports updates, even though the cursor is based on a multiple table join — some columns of which are not updateable.</span></span>
