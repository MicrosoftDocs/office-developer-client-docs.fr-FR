---
title: SkipLine, méthode (ADO)
TOCTitle: SkipLine method (ADO)
ms:assetid: 419c24c3-6b84-eed0-5884-f2dcd485dc3d
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249187(v=office.15)
ms:contentKeyID: 48544456
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: 5e49b8379f078ad698a4a0040de0eb2e4429fd34
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32314559"
---
# <a name="skipline-method-ado"></a>SkipLine, méthode (ADO)


**S’applique à** : Access 2013, Office 2013

Saute une ligne complète lors de la lecture d'un flux de texte.

## <a name="syntax"></a>Syntaxe

*Stream*. SkipLine

## <a name="remarks"></a>Remarques

Tous les caractères, y compris le séparateur de ligne suivante, sont ignorés. Par défaut, [LineSeparator](lineseparator-property-ado.md) prend la valeur **adCRLF**. Si vous tentez de revenir au-delà de [EOS](eos-property-ado.md), la position actuelle reste simplement au niveau de **EOS**.

La méthode **SkipLine** est utilisée avec des flux de texte ([Type](type-property-ado-stream.md) prend la valeur **adTypeText**).

