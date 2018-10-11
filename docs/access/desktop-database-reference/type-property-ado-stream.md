---
title: Type, propriété (objet Stream ADO)
TOCTitle: Type Property (ADO Stream)
ms:assetid: 43872c74-51bf-47ae-6bdc-55d25b0dc84a
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249203(v=office.15)
ms:contentKeyID: 48544505
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 2a7eaf97e61ffb1abfed3104644936867c325641
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470185"
---
# <a name="type-property-ado-stream"></a><span data-ttu-id="681b4-102">Type, propriété (objet Stream ADO)</span><span class="sxs-lookup"><span data-stu-id="681b4-102">Type Property (ADO Stream)</span></span>


<span data-ttu-id="681b4-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="681b4-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="681b4-104">Indique le type de données contenu dans l'objet [Stream](stream-object-ado.md) (binaire ou texte).</span><span class="sxs-lookup"><span data-stu-id="681b4-104">Indicates the type of data contained in the [Stream](stream-object-ado.md) (binary or text).</span></span>

## <a name="settings-and-return-values"></a><span data-ttu-id="681b4-105">Paramètres et valeurs renvoyées</span><span class="sxs-lookup"><span data-stu-id="681b4-105">Settings and Return Values</span></span>

<span data-ttu-id="681b4-p101">Définit ou renvoie une valeur [StreamTypeEnum](streamtypeenum.md) qui spécifie le type de données contenu dans l'objet **Stream**. La valeur par défaut est **adTypeText**. Toutefois, si ce sont des données binaires qui sont initialement écrites dans un nouvel objet **Stream** vide, la valeur de la propriété **Type** devient **adTypeBinary**.</span><span class="sxs-lookup"><span data-stu-id="681b4-p101">Sets or returns a [StreamTypeEnum](streamtypeenum.md) value that specifies the type of data contained in the **Stream** object. The default value is **adTypeText**. However, if binary data is initially written to a new, empty **Stream**, the **Type** will be changed to **adTypeBinary**.</span></span>

## <a name="remarks"></a><span data-ttu-id="681b4-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="681b4-109">Remarks</span></span>

<span data-ttu-id="681b4-110">La propriété **Type** n’est accessible en lecture et en écriture que lorsque la position actuelle est au début de l’objet **Stream** (la valeur de la propriété [Position](position-property-ado.md) est 0) ; elle est en lecture seule pour toute autre position.</span><span class="sxs-lookup"><span data-stu-id="681b4-110">The **Type** property is read/write only when the current position is at the beginning of the **Stream** ([Position](position-property-ado.md) is 0), and read-only at any other position.</span></span>

<span data-ttu-id="681b4-p102">La propriété **Type** détermine les méthodes devant être utilisées pour effectuer des lectures et des écritures sur l'objet **Stream**. Pour les objets **Stream** textuels, utilisez [ReadText](readtext-method-ado.md) et [Writetext](writetext-method-ado.md). Pour les objets **Stream** binaires, utilisez [Read](read-method-ado.md) et [Write](write-method-ado.md).</span><span class="sxs-lookup"><span data-stu-id="681b4-p102">The **Type** property determines which methods should be used for reading and writing the **Stream**. For text **Streams**, use [ReadText](readtext-method-ado.md) and [WriteText](writetext-method-ado.md). For binary **Streams**, use [Read](read-method-ado.md) and [Write](write-method-ado.md).</span></span>
