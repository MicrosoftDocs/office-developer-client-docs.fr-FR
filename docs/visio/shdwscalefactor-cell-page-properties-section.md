---
title: ShdwScaleFactor, cellule (section Page Properties)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- vis_sdr.chm60083
localization_priority: Normal
ms.assetid: 10979706-6dfe-5241-e862-3f94716d14fa
description: Indique le pourcentage d'agrandissement ou de réduction de la forme d'une ombre.
ms.openlocfilehash: 9175e9a1148779524fdce96ff18eac22fe8dd421
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33429004"
---
# <a name="shdwscalefactor-cell-page-properties-section"></a>ShdwScaleFactor, cellule (section Page Properties)

Indique le pourcentage d'agrandissement ou de réduction de la forme d'une ombre. 
  
## <a name="remarks"></a>Remarques

Chaque ombre a un axe d'ombre portée, qui est un point sur l'ombre qui correspond à l'axe de la forme. Par exemple, si l'axe d'une forme est au centre de cette dernière, l'axe d'ombre portée est le point central de l'ombre. Lors de l’application d’une échelle à des ombres simples, le grossissement est centré à l’emplacement de la broche ombrée ; lors de l’application d’une échelle aux ombres obliques, le grossissement est appliqué dans le sens oblique. 
  
 Ce pourcentage est utilisé lorsque le type d’ombre d’une forme est définie sur Page Default (la cellule ShapeShdwType est égale à ** visFSTPageDefault ** ). 
  
Pour définir ce comportement pour une forme individuelle, utilisez la cellule ShapeShdwScaleFactor dans la section Fill Format.
  
Cette valeur correspond à celle de la zone **Facteur d’agrandissement** sous l’onglet **Ombres** de la boîte de dialogue **Mise en page** (sous l’onglet **Création**, cliquez sur **Mise en page**). 
  
Pour obtenir une référence à la cellule ShdwScaleFactor par un nom dans une autre formule ou dans un programme en faisant appel à la propriété **CellsU**, utilisez : 
  
|||
|:-----|:-----|
| Nom de cellule :  <br/> | ShdwScaleFactor  <br/> |
   
Pour obtenir une référence à la cellule ShdwScaleFactor par index dans un programme, utilisez la propriété **CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
| Index de la section :  <br/> |**visSectionObject** <br/> |
| Index de la ligne :  <br/> |**visRowPage** <br/> |
| Index de la cellule :  <br/> |**visPageShdwScaleFactor** <br/> |
   

