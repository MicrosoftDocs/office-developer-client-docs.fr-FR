---
title: ASIN, fonction
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- Vis_DSS.chm82251395
localization_priority: Normal
ms.assetid: 7d917be4-65b1-002f-48cc-6d81916a1157
description: Renvoie l’arcsine d’un nombre, par exemple, l’angle dont le sinus est le nombre .
ms.openlocfilehash: e66ed9ab3d01ac79bceb18f5c793afc928e5e4b4
ms.sourcegitcommit: 939bd9686ba41a8f94b82e004ed84b9054d9c7cf
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2020
ms.locfileid: "48293505"
---
# <a name="asin-function"></a>Fonction ASIN

Renvoie l’arcsine d’un nombre, par exemple, l’angle dont le sinus est  *le nombre*  . 
  
## <a name="syntax"></a>Syntaxe

ASIN(***nombre*** ) 
  
### <a name="parameters"></a>Paramètres

|**Nom**|**Requis/Facultatif**|**Type de données**|**Description**|
|:-----|:-----|:-----|:-----|
| _number_ <br/> |Obligatoire  <br/> |**Numérique** <br/> |Sinus de l’angle.  <br/> |
   
## <a name="remarks"></a>Remarques

La valeur d’entrée doit être dans la plage -1 <=  *nombre*  <= 1 ou une valeur #NUM! est renvoyée. L’angle obtenu se trouve dans la plage -PI/2 <=  *angle*  <= PI/2 radians (-90 <=  *angle*  <= 90 degrés). 
  
## <a name="example"></a>Exemple

ASIN(1)
  
Renvoie 90 deg
  

