---
title: PrintGrid, cellule (section Print Properties)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- Vis_DSS.chm1033794
localization_priority: Normal
ms.assetid: 0504ff7f-2274-7ae3-1f4b-a3d890dbd79a
description: Indique si la grille sera imprimée lors de l'impression d'une page de document.
ms.openlocfilehash: 9b98999cd02fa6a47ec8564bbd7337ecf8637306
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33433205"
---
# <a name="printgrid-cell-print-properties-section"></a>PrintGrid, cellule (section Print Properties)

Indique si la grille sera imprimée lors de l'impression d'une page de document.
  
|**Valeur**|**Description**|
|:-----|:-----|
|TRUE  <br/> |Affiche la grille lors de l'impression de cette page.  <br/> |
|FALSE  <br/> |N'affiche pas la grille lors de l'impression de cette page (valeur par défaut).  <br/> |
   
## <a name="remarks"></a>Remarques

Cette valeur correspond à la case à cocher **Quadrillage** sous l’onglet **Configuration** de la boîte de dialogue **Mise en page** (sous l’onglet **Création**, cliquez sur **Mise en page**). Exception faite de la couleur (la version imprimée est grise), la grille imprimée est identique à celle que vous voyez dans la fenêtre de dessin Microsoft Office Visio. 
  
Vous pouvez choisir d’imprimer la grille page par page. Le style de grille peut également être défini page par page dans la  boîte de  dialogue Grille de règle (sous l’onglet Affichage, cliquez sur Afficher la flèche) lorsqu’une page est active. **&amp;** 
  
Pour obtenir une référence à la cellule PrintGrid par un nom à partir d'une autre formule ou d'un programme en faisant appel à la propriété **CellsU**, utilisez : 
  
|||
|:-----|:-----|
|Nom de la cellule :  <br/> |PrintGrid  <br/> |
   
Pour obtenir une référence à la cellule PrintGrid à l'aide d'un index à partir d'un programme, utilisez la propriété **CellsSRC** avec les arguments suivants : 
  
|||
|:-----|:-----|
|Index de la section :  <br/> |**visSectionObject** <br/> |
|Index de la ligne :  <br/> |**visRowPrintProperties** <br/> |
|Index de la cellule :  <br/> |**visPrintPropertiesPrintGrid** <br/> |
   

