---
title: BOUNDINGBOXDIST, fonction
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: 8a2490f2-48c4-5df3-a3b3-40e8e0c80479
description: Renvoie la mesure de la partie spécifiée du cadre englobant de la forme.
ms.openlocfilehash: a62d5b82c310e7b99e16b70982b75a68172b7fd8
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33428276"
---
# <a name="boundingboxdist-function"></a>Fonction BOUNDINGBOXDIST

Renvoie la mesure de la partie spécifiée du cadre englobant de la forme. 
  
## <a name="version-information"></a>Informations de version

Version ajoutée : Visio 2010
 
  
## <a name="syntax"></a>Syntaxe

BOUNDINGBOXDIST(** *Index* ** ) 
  
### <a name="parameters"></a>Paramètres

|**Nom**|**Requis/Facultatif**|**Type de données**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_ <br/> |Obligatoire  <br/> |**Number** <br/> |Partie du cadre de limite de la forme à mesurer et à renvoyer. Les valeurs possibles, reportez-vous à la section Remarques.  <br/> |
   
### <a name="return-value"></a>Valeur renvoyée

 **Number**
  
## <a name="remarks"></a>Remarques

 *Index*  peut être l’une des valeurs suivantes. 
  
|**Élément**|**Valeur**|
|:-----|:-----|
|Largeur  <br/> |0  <br/> |
|Hauteur  <br/> |1  <br/> |
|Bord gauche à l’axe de la forme  <br/> |2  <br/> |
|Axe de la forme au bord droit  <br/> |3  <br/> |
|Axe de la forme au bord supérieur  <br/> |4   <br/> |
|Bord inférieur à l’axe de la forme  <br/> |5   <br/> |
|Centre du cadre englobant à AxeX  <br/> |6   <br/> |
|Centre du cadre englobant à AxeY  <br/> |7   <br/> |
   

