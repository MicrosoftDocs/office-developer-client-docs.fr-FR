---
title: PageWidth, cellule (section Page Properties)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- Vis_DSS.chm82251372
localization_priority: Normal
ms.assetid: b98c5bf3-10c8-7299-2836-3906d6a9135d
description: Détermine la largeur de la page imprimée en unités de dessin.
ms.openlocfilehash: 6d887cb4335d2725101db54ba2b1483ccf01cff4
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33434269"
---
# <a name="pagewidth-cell-page-properties-section"></a>PageWidth, cellule (section Page Properties)

Détermine la largeur de la page imprimée en unités de dessin.
  
## <a name="remarks"></a>Remarques

Vous pouvez également définir la largeur de  page sous l’onglet  Taille de page de la boîte de dialogue Mise en page (sous l’onglet Création, cliquez sur la flèche Mise en **page)** ou en re dimensionner manuellement la page avec la souris.  Pour ce faire, faites glisser le bord de la page en maintenant la touche Ctrl enfoncée. 
  
Pour obtenir une référence à la cellule PageWidth par un nom à partir d'une autre formule ou d'un programme en faisant appel à la propriété **CellsU**, utilisez : 
  
|||
|:-----|:-----|
|Nom de la cellule :  <br/> |PageWidth  <br/> |
   
Pour obtenir une référence à la cellule PageWidth à l'aide d'un index à partir d'un programme, utilisez la propriété **CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
|Index de la section :  <br/> |**visSectionObject** <br/> |
|Index de la ligne :  <br/> |**visRowPage** <br/> |
|Index de la cellule :  <br/> |**visPageWidth** <br/> |
   

