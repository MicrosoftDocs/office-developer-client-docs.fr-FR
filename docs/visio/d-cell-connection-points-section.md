---
title: D, cellule (section Connection Points)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- vis_sdr.chm205
localization_priority: Normal
ms.assetid: 28b18e8d-fecf-a798-813e-c1a310002244
description: Cellule de montage utilisable pour entrer ou tester des formules.
ms.openlocfilehash: e7bd61b8bc7a1a3b765af738681d958e2c83ba05
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33408172"
---
# <a name="d-cell-connection-points-section"></a>D, cellule (section Connection Points)

Cellule de montage utilisable pour entrer ou tester des formules.
  
## <a name="remarks"></a>Remarques

Pour accéder à la cellule D, cliquez avec le bouton droit de la souris sur une ligne, puis cliquez sur **Modifier le type de ligne** dans le menu contextuel. 
  
Pour obtenir une référence à la cellule D par un nom à partir d'une autre formule ou d'un programme en faisant appel à la propriété **CellsU**, utilisez : 
  
|||
|:-----|:-----|
| Nom de la cellule :  <br/> | Connections.D[  *i*  ] où  *i*  = <1>, 2, 3...  <br/> |
   
Pour obtenir une référence à la cellule D à l'aide d'un index à partir d'un programme, utilisez la propriété **CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
| Index de la section :  <br/> |**visSectionConnectionPts** <br/> |
| Index de la ligne :  <br/> |**visRowConnectionPts**  +   *i* où *i* = 0, 1, 2...  <br/> |
| Index de la cellule :  <br/> |**visCnnctD** <br/> |
   

