---
title: LineWeight, cellule (section Line Format)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- Vis_DSS.chm585
localization_priority: Normal
ms.assetid: 16b0e293-eeef-34b4-aeb0-4472815dd543
description: Détermine l’épaisseur d’un trait. Pour définir l’épaisseur, entrez un nombre accompagné d’une unité de mesure valide.
ms.openlocfilehash: 654a93f939226bedab2e40ab591dad0e3f872267
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33412246"
---
# <a name="lineweight-cell-line-format-section"></a>LineWeight, cellule (section Line Format)

Détermine l’épaisseur d’un trait. Pour définir l’épaisseur, entrez un nombre accompagné d’une unité de mesure valide.
  
## <a name="remarks"></a>Remarques

Vous pouvez également définir la valeur de LineWeight dans la boîte de dialogue **Trait** (sous l’onglet **Trait**, dans le groupe **Forme**, cliquez sur **Trait**, pointez sur **Épaisseur**, puis cliquez sur **Autres traits**).
  
Si l’unité de mesure n’est pas entrée, l’unité de mesure du texte  spécifié dans la boîte de dialogue **Options Visio** est utilisée (cliquez sur l’onglet Fichier, puis sur **Options).** L’épaisseur de trait est indépendante de l’échelle du dessin. Si le dessin est mis à l’échelle, l’épaisseur de trait ne change pas. 
  
Pour obtenir une référence à la cellule LineWeight par un nom à partir d’une autre formule ou d’un programme en faisant appel à la propriété **CellsU**, utilisez : 
  
|||
|:-----|:-----|
| Nom de la cellule :  <br/> | LineWeight  <br/> |
   
Pour obtenir une référence à la cellule LineWeight à l’aide d’un index à partir d’un programme, utilisez la propriété **CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
| Index de la section :  <br/> |**visSectionObject** <br/> |
| Index de la ligne :  <br/> |**visRowLine** <br/> |
| Index de la cellule :  <br/> |**visLineWeight** <br/> |
   

