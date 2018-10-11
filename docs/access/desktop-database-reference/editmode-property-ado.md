---
title: EditMode, propriété (ADO)
TOCTitle: EditMode Property (ADO)
ms:assetid: 28ca8f14-abee-ad20-9c16-11bb36b487e4
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249045(v=office.15)
ms:contentKeyID: 48543867
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: a6c8377e0fa0fc2db1ec2d376ee3c8b9f16e8c99
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469793"
---
# <a name="editmode-property-ado"></a><span data-ttu-id="c6c87-102">EditMode, propriété (ADO)</span><span class="sxs-lookup"><span data-stu-id="c6c87-102">EditMode Property (ADO)</span></span>


<span data-ttu-id="c6c87-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="c6c87-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="c6c87-104">Indique l'état d'édition de l'enregistrement actuel.</span><span class="sxs-lookup"><span data-stu-id="c6c87-104">Indicates the editing status of the current record.</span></span>

## <a name="return-value"></a><span data-ttu-id="c6c87-105">Valeur renvoyée</span><span class="sxs-lookup"><span data-stu-id="c6c87-105">Return Value</span></span>

<span data-ttu-id="c6c87-106">Renvoie la valeur [EditModeEnum](editmodeenum.md).</span><span class="sxs-lookup"><span data-stu-id="c6c87-106">Returns an [EditModeEnum](editmodeenum.md) value.</span></span>

## <a name="remarks"></a><span data-ttu-id="c6c87-107">Remarques</span><span class="sxs-lookup"><span data-stu-id="c6c87-107">Remarks</span></span>

<span data-ttu-id="c6c87-p101">ADO gère un tampon d'édition associé à l'enregistrement actuel. Cette propriété indique si des modifications ont été apportées à ce tampon ou si un nouvel enregistrement a été créé. Utilisez la propriété **EditMode** pour déterminer l'état d'édition de l'enregistrement actuel. Vous pouvez tester jusqu'à des modifications si un processus d'édition a été interrompu et détermine si vous devez utiliser la méthode [Update](update-method-ado.md) ou [CancelUpdate](cancelupdate-method-ado.md).</span><span class="sxs-lookup"><span data-stu-id="c6c87-p101">ADO maintains an editing buffer associated with the current record. This property indicates whether changes have been made to this buffer, or whether a new record has been created. Use the **EditMode** property to determine the editing status of the current record. You can test for pending changes if an editing process has been interrupted and determine whether you need to use the [Update](update-method-ado.md) or [CancelUpdate](cancelupdate-method-ado.md) method.</span></span>

<span data-ttu-id="c6c87-112">Consultez la méthode [AddNew](addnew-method-ado.md) pour une description plus détaillée de la propriété **EditMode** dans des conditions d'édition différentes.</span><span class="sxs-lookup"><span data-stu-id="c6c87-112">See the [AddNew](addnew-method-ado.md) method for a more detailed description of the **EditMode** property under different editing conditions.</span></span>

<span data-ttu-id="c6c87-113">Lorsqu’un appel à [Supprimer](delete-method-ado-recordset.md) ne supprime pas correctement l’enregistrement ou des enregistrements dans les données source (en raison de violations d’intégrité référentielle, par exemple), l' [objet Recordset](recordset-object-ado.md) reste en mode édition (**EditMode** = \*\*adEditInProgress \*\*).</span><span class="sxs-lookup"><span data-stu-id="c6c87-113">When a call to [Delete](delete-method-ado-recordset.md) does not successfully delete the record or records in the data source (due to referential integrity violations, for example), the [Recordset](recordset-object-ado.md) will remain in edit mode (**EditMode** = **adEditInProgress**).</span></span> <span data-ttu-id="c6c87-114">Il faut donc appeler **CancelUpdate** avant de déplacer l'enregistrement actuel (avec [Move](move-method-ado.md), [NextRecordset](nextrecordset-method-ado.md) ou [Close](close-method-ado.md), par exemple).</span><span class="sxs-lookup"><span data-stu-id="c6c87-114">This means that **CancelUpdate** must be called before moving off the current record (with [Move](move-method-ado.md), [NextRecordset](nextrecordset-method-ado.md), or [Close](close-method-ado.md), for example).</span></span>


> [!NOTE]
> <P><span data-ttu-id="c6c87-p103">[!REMARQUE] <STRONG>EditMode</STRONG> ne peut renvoyer une valeur valide que s'il existe un enregistrement actuel. <STRONG>EditMode</STRONG> renvoie une erreur si la valeur de <A href="bof-eof-properties-ado.md">BOF or EOF</A> est True ou si l'enregistrement actuel a été supprimé.</span><span class="sxs-lookup"><span data-stu-id="c6c87-p103"><STRONG>EditMode</STRONG> can return a valid value only if there is a current record. <STRONG>EditMode</STRONG> will return an error if <A href="bof-eof-properties-ado.md">BOF or EOF</A> is true, or if the current record has been deleted.</span></span></P>

