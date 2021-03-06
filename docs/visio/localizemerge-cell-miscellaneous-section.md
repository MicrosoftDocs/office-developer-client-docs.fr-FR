---
title: LocalizeMerge, cellule (section Miscellaneous)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- Vis_DSS.chm1033773
localization_priority: Normal
ms.assetid: 734d4415-05dd-4c4d-763e-e035fa56dcec
description: Détermine si les formes sont localisées lors de leur copie d'un document à l'autre.
ms.openlocfilehash: ddd6041ec6531652deb38a0c16be2c741bac91a6
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33405673"
---
# <a name="localizemerge-cell-miscellaneous-section"></a>LocalizeMerge, cellule (section Miscellaneous)

Détermine si les formes sont localisées lors de leur copie d'un document à l'autre.
  
|**Valeur**|**Description**|
|:-----|:-----|
| TRUE  <br/> | Localise une forme dans la langue du document de destination.  <br/> |
| FALSE  <br/> | Ne localisez pas une forme en fonction de la langue du document de destination (valeur par défaut).  <br/> |
   
## <a name="remarks"></a>Remarques

Pour obtenir une référence à la cellule LocalizeMerge par un nom à partir d'une autre formule ou d'un programme en faisant appel à la propriété **CellsU**, utilisez : 
  
|||
|:-----|:-----|
| Nom de la cellule :  <br/> | LocalizeMerge  <br/> |
   
Pour obtenir une référence à la cellule LocalizeMerge à l'aide d'un index à partir d'un programme, utilisez la propriété **CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
| Index de la section :  <br/> |**visSectionObject** <br/> |
| Index de la ligne :  <br/> |**visRowMisc** <br/> |
| Index de la cellule :  <br/> |**visObjLocalizeMerge** <br/> |
   

