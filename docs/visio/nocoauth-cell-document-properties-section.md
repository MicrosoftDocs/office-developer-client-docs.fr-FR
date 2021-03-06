---
title: NoCoauth Cell (Document Properties Section)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: 6f2095c9-ce09-48f7-b160-c9822d96a96c
description: Définit si un document stocké sur un serveur microsoft SharePoint 2013 ou un Microsoft OneDrive peut être modifié simultanément par plusieurs auteurs dans une session de co-édition.
ms.openlocfilehash: a76e2d3b2c3cf6e99e37596b016f448b0be56fd3
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33420513"
---
# <a name="nocoauth-cell-document-properties-section"></a>NoCoauth Cell (Document Properties Section)

Définit si un document stocké sur un serveur microsoft SharePoint 2013 ou un Microsoft OneDrive peut être modifié simultanément par plusieurs auteurs dans une session de co-édition.
  
|**Valeur**|**Description**|
|:-----|:-----|
|TRUE  <br/> |Le document ne peut pas être co-écrit et est verrouillé pour modification à l’ouverture.  <br/> |
|FALSE  <br/> |Le document peut être co-écrit.  <br/> |
   
## <a name="remarks"></a>Remarques

Pour obtenir une référence à la cellule **NoCoauth** par un nom à partir d’une autre formule, de l’attribut **N** d’un élément **Cell** ou d’un programme en faisant appel à la propriété **CellsU,** utilisez : 
  
|||
|:-----|:-----|
| Nom de cellule :  <br/> | NoCoauth  <br/> |
   
Pour obtenir une référence à la **cellule NoCoauth** à l’aide d’un index à partir d’un programme, utilisez la propriété **CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
| Index de la section :  <br/> |**visSectionObject** <br/> |
| Index de la ligne :  <br/> |**visRowDoc** <br/> |
| Index de la cellule :  <br/> |**visDocNoCoauth** <br/> |
   

