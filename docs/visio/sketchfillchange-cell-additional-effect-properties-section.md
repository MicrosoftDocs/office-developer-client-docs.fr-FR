---
title: SketchFillChange Cell (Additional Effect Properties Section)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: 939f8f90-dee5-4175-b32a-e2964eb40681
description: Détermine la quantité de randomisation du remplissage de la forme à partir de la géométrie de la forme lors de l’utilisation d’un effet de croquis, sous la forme d’un pourcentage de la longueur d’une section. Si la valeur de la cellule SketchFillChange est définie sur 0 %, la géométrie de limite du remplissage d’une forme correspond à la géométrie de la forme. Si la valeur est 100 %, la géométrie de limite du remplissage de la forme ne suit pas la géométrie de la forme.
ms.openlocfilehash: 8726e9dd6ca6257fb8dbbbef3dce1d4ec344e28b
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33408074"
---
# <a name="sketchfillchange-cell-additional-effect-properties-section"></a>SketchFillChange Cell (Additional Effect Properties Section)

Détermine la quantité de randomisation du remplissage de la forme à partir de la géométrie de la forme lors de l’utilisation d’un effet de croquis, sous la forme d’un pourcentage de la longueur d’une section. Si la valeur de la **cellule SketchFillChange** est définie sur 0 %, la géométrie de limite du remplissage d’une forme correspond à la géométrie de la forme. Si la valeur est 100 %, la géométrie de limite du remplissage de la forme ne suit pas la géométrie de la forme. 
  
## <a name="remarks"></a>Remarques

Pour de meilleurs résultats, la plage idéale de valeurs pour la cellule **SketchFillChange** est entre 15 et 50 %. Une valeur inférieure à 15 % est à peine perceptible . une valeur supérieure à 50 % est de plus en plus aléatoire. 
  
Pour obtenir une référence à la cellule **SketchFillChange** par un nom à partir d’une autre formule, de l’attribut **N** d’un élément **Cell** ou d’un programme en faisant appel à la propriété **CellsU,** utilisez : 
  
|||
|:-----|:-----|
| Nom de cellule :  <br/> | SketchFillChange  <br/> |
   
Pour obtenir une référence à la **cellule SketchFillChange** à l’aide d’un index à partir d’un programme, utilisez la **propriété CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
| Index de la section :  <br/> |**visSectionObject** <br/> |
| Index de la ligne :  <br/> |**visRowOtherEffectProperties** <br/> |
| Index de la cellule :  <br/> |**visSketchFillChange** <br/> |
   

