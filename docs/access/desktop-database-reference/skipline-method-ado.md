---
title: SkipLine, méthode (ADO)
TOCTitle: SkipLine Method (ADO)
ms:assetid: 419c24c3-6b84-eed0-5884-f2dcd485dc3d
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249187(v=office.15)
ms:contentKeyID: 48544456
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 28cb222a3ed0e5497e03efbb7394081c96b4d4ef
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470810"
---
# <a name="skipline-method-ado"></a><span data-ttu-id="632cc-102">SkipLine, méthode (ADO)</span><span class="sxs-lookup"><span data-stu-id="632cc-102">SkipLine Method (ADO)</span></span>


<span data-ttu-id="632cc-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="632cc-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="632cc-104">Saute une ligne complète lors de la lecture d'un flux de texte.</span><span class="sxs-lookup"><span data-stu-id="632cc-104">Skips one entire line when reading a text stream.</span></span>

## <a name="syntax"></a><span data-ttu-id="632cc-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="632cc-105">Syntax</span></span>

<span data-ttu-id="632cc-106">*Flux de données*. SkipLine</span><span class="sxs-lookup"><span data-stu-id="632cc-106">*Stream*.SkipLine</span></span>

## <a name="remarks"></a><span data-ttu-id="632cc-107">Notes</span><span class="sxs-lookup"><span data-stu-id="632cc-107">Remarks</span></span>

<span data-ttu-id="632cc-p101">Tous les caractères, y compris le séparateur de ligne suivante, sont ignorés. Par défaut, [LineSeparator](lineseparator-property-ado.md) prend la valeur **adCRLF**. Si vous tentez de revenir au-delà de [EOS](eos-property-ado.md), la position actuelle reste simplement au niveau de **EOS**.</span><span class="sxs-lookup"><span data-stu-id="632cc-p101">All characters up to, and including the next line separator, are skipped. By default, the [LineSeparator](lineseparator-property-ado.md) is **adCRLF**. If you attempt to skip past [EOS](eos-property-ado.md), the current position will simply remain at **EOS**.</span></span>

<span data-ttu-id="632cc-111">La méthode **SkipLine** est utilisée avec des flux de texte ([Type](type-property-ado-stream.md) prend la valeur **adTypeText**).</span><span class="sxs-lookup"><span data-stu-id="632cc-111">The **SkipLine** method is used with text streams ([Type](type-property-ado-stream.md) is **adTypeText**).</span></span>
