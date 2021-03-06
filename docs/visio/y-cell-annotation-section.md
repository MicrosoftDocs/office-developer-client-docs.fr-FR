---
title: Y, cellule (section Annotation)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- vis_sdr.chm60095
localization_priority: Normal
ms.assetid: 527a4615-2013-a4b4-81cd-7f5d71c1803c
description: Coordonnée y du marqueur de commentaire en coordonnées de page.
ms.openlocfilehash: 48a37c261078cd1000331920b33549cee2c1da03
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33429200"
---
# <a name="y-cell-annotation-section"></a>Y, cellule (section Annotation)

Coordonnée  *y*  du marqueur de commentaire en coordonnées de page. 
  
> [!NOTE]
> Cette cellule est utilisée pour suivre les commentaires uniquement lors de l’ouverture d’un fichier .vsd dans Microsoft Visio 2013 ou lors de l’enregistrement d’un fichier .vsdx au format de fichier .vsd. Il n’est pas utilisé pour le suivi des commentaires dans les documents .vsdx Visio 2013. 
  
## <a name="remarks"></a>Remarques

Pour obtenir une référence à la cellule Y à l'aide d'un nom à partir d'une autre formule ou programme en faisant appel à la propriété **CellsU**, utilisez : 
  
|||
|:-----|:-----|
| Nom de cellule :  <br/> | Annotation.Y [  *i*  ] où  *i*  = <1>, 2, 3...  <br/> |
   
Pour obtenir une référence à la cellule Y à l'aide d'un index à partir d'un programme, utilisez la propriété **CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
| Index de la section :  <br/> |**visSectionAnnotation** <br/> |
| Index de la ligne :  <br/> |**visRowAnnotation**  +   *i* où *i* = 0, 1, 2...  <br/> |
| Index de la cellule :  <br/> |**visAnnotationY** <br/> |
   

