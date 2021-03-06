---
title: IS [NOT] NULL (application web personnalisée Access)
manager: kelbow
ms.date: 09/05/2017
ms.audience: Developer
ms.topic: reference
ms.assetid: b941a0c7-9753-4920-bb6d-cbba94ba9422
description: Détermine si une expression spécifiée est NULL.
localization_priority: Priority
ms.openlocfilehash: fe6a0fe4f182a1385304b783e7cfaaf515f732d4
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32311136"
---
# <a name="is-not-null-access-custom-web-app"></a>IS [NOT] NULL (application web personnalisée Access)

Détermine si une expression spécifiée est NULL.
  
> [!IMPORTANT]
> Microsoft ne recommande plus la création et l'utilisation d'applications web Access dans SharePoint. En guise d'alternative, vous pouvez utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/en-us/) pour générer des solutions d'entreprise sans code pour le web et les appareils mobiles. 
  
## <a name="syntax"></a>Syntaxe

 *expression* **IS** [  *NOT*  ] **NULL**
  
Le prédicat **IS [NOT] NULL** contient les arguments suivants. 
  
|||
|:-----|:-----|
| *expression*  <br/> |Toute expression valide.  <br/> |
| *NOT*  <br/> |Spécifie que le résultat booléen doit être réfuté. Le prédicat inverse ses valeurs renvoyées, renvoyant TRUE si la valeur n’est pas NULL et FALSE si la valeur est NULL.  <br/> |
   
## <a name="remarks"></a>Remarques

Si la valeur de l’*expression* est NULL, IS NULL renvoie TRUE ; dans le cas contraire, elle renvoie FALSE. 
  
Si la valeur de l’expression est NULL, IS NOT NULL renvoie FALSE ; dans le cas contraire, elle renvoie TRUE.
  

