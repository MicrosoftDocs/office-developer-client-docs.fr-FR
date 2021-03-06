---
title: PlaceDepth, cellule (section Page Layout)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- Vis_DSS.chm1230
localization_priority: Normal
ms.assetid: 02c139db-fe67-f550-1d07-8c8a9a4fb427
description: Détermine la méthode utilisée pour analyser le dessin avant la création de la mise en page et définit le type de mise en page.
ms.openlocfilehash: 463c7dad39955161538aa89d1482685189bf7fdc
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33432036"
---
# <a name="placedepth-cell-page-layout-section"></a>PlaceDepth, cellule (section Page Layout)

Détermine la méthode utilisée pour analyser le dessin avant la création de la mise en page et définit le type de mise en page.
  
|**Valeur**|**Profondeur de placement pour les mises en page verticales et horizontales**|**Constante d'automation**|
|:-----|:-----|:-----|
| 0  <br/> | Valeur par défaut de la page  <br/> |**visPLOPlaceDepthDefault** <br/> |
| 1  <br/> | Moyenne  <br/> |**visPLOPlaceDepthMedium** <br/> |
| 2  <br/> | Deep  <br/> |**visPLOPlaceDepthDeep** <br/> |
| 3  <br/> | Superficiel  <br/> |**visPLOPlaceDepthShallow** <br/> |
   
## <a name="remarks"></a>Remarques

Pour obtenir une référence à la cellule PlaceDepth par un nom à partir d'une autre formule ou d'un programme en faisant appel à la propriété **CellsU**, utilisez : 
  
|||
|:-----|:-----|
| Nom de la cellule :  <br/> | PlaceDepth  <br/> |
   
Pour obtenir une référence à la cellule PlaceDepth à l'aide d'un index à partir d'un programme, utilisez la propriété **CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
| Index de la section :  <br/> |**visSectionObject** <br/> |
| Index de la ligne :  <br/> |**visRowPageLayout** <br/> |
| Index de la cellule :  <br/> |**visPLOPlaceDepth** <br/> |
   

