---
title: Precision, propriété (ADO)
TOCTitle: Precision Property (ADO)
ms:assetid: c9d54d78-d5a5-caf8-d635-259d1fcc0595
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249983(v=office.15)
ms:contentKeyID: 48547685
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: a9e98227e98b4f731cd4d361ff6a2c5394058e86
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472508"
---
# <a name="precision-property-ado"></a><span data-ttu-id="23f17-102">Precision, propriété (ADO)</span><span class="sxs-lookup"><span data-stu-id="23f17-102">Precision Property (ADO)</span></span>


<span data-ttu-id="23f17-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="23f17-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="23f17-104">Indique le degré de précision des valeurs numériques d'un objet [Parameter](parameter-object-ado.md) ou d'un objet [Field](field-object-ado.md) numérique.</span><span class="sxs-lookup"><span data-stu-id="23f17-104">Indicates the degree of precision for numeric values in a [Parameter](parameter-object-ado.md) object or for numeric [Field](field-object-ado.md) objects.</span></span>

## <a name="settings-and-return-values"></a><span data-ttu-id="23f17-105">Paramètres et valeurs renvoyées</span><span class="sxs-lookup"><span data-stu-id="23f17-105">Settings and Return Values</span></span>

<span data-ttu-id="23f17-106">Définit ou renvoie une valeur de type **Byte** qui indique le nombre maximal de chiffres utilisés pour représenter les valeurs.</span><span class="sxs-lookup"><span data-stu-id="23f17-106">Sets or returns a **Byte** value that indicates the maximum number of digits used to represent values.</span></span>

## <a name="remarks"></a><span data-ttu-id="23f17-107">Remarques</span><span class="sxs-lookup"><span data-stu-id="23f17-107">Remarks</span></span>

<span data-ttu-id="23f17-108">Utilisez la propriété **Précision** pour déterminer le nombre maximal de chiffres utilisés pour représenter les valeurs d'un objet **Parameter** ou **Field** numérique.</span><span class="sxs-lookup"><span data-stu-id="23f17-108">Use the **Precision** property to determine the maximum number of digits used to represent values for a numeric **Parameter** or **Field** object.</span></span>

<span data-ttu-id="23f17-109">Sur un objet **Parameter** cette valeur est accessible lecture/écriture.</span><span class="sxs-lookup"><span data-stu-id="23f17-109">The value is read/write on a **Parameter** object.</span></span>

<span data-ttu-id="23f17-p101">Pour un objet **Field**, **Precision** est généralement en lecture seule. Toutefois, pour les objets **Field** récemment ajoutés à la collection [Fields](fields-collection-ado.md) d'un objet [Record](record-object-ado.md), **Precision** est exceptionnellement accessible en lecture et en écriture après la définition de la propriété [Value](value-property-ado.md) de l'objet **Field** et une fois que le fournisseur de données a ajouté le nouvel objet **Field** en appelant la méthode [Update](update-method-ado.md) de la collection **Fields**.</span><span class="sxs-lookup"><span data-stu-id="23f17-p101">For a **Field** object, **Precision** is normally read-only. However, for new **Field** objects that have been appended to the [Fields](fields-collection-ado.md) collection of a [Record](record-object-ado.md), **Precision** is read/write only after the [Value](value-property-ado.md) property for the **Field** has been specified and the data provider has successfully added the new **Field** by calling the [Update](update-method-ado.md) method of the **Fields** collection.</span></span>
