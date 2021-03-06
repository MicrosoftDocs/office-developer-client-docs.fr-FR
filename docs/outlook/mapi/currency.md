---
title: CURRENCY
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- CURRENCY
api_type:
- COM
ms.assetid: cffc05a0-95e4-4b9f-bf8f-c4272a75afa8
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: dccb6b19af72d0f748a3a513b7f3d78904ebc789
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33431119"
---
# <a name="currency"></a>CURRENCY

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient un integer signé 64 bits représentant une valeur monétaire. 
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapidefs.h  <br/> |
   
```cpp
typedef struct tagCY
{
  unsigned long Lo;
  long Hi;
} CURRENCY;

```

## <a name="members"></a>Members

 **Lo**
  
> 32 bits de bas ordre de la valeur monétaire. 
    
 **Salut**
  
> 32 bits de haut ordre de la valeur monétaire.
    
## <a name="remarks"></a>Remarques

La structure **CURRENCY** est une représentation entière à l’échelle d’un nombre décimal à quatre chiffres à droite de la virgule décimale. Par exemple, une valeur stockée de 327 500 doit être interprétée comme représentant une valeur monétaire de 32,7500. 
  
La structure **CURRENCY** est utilisée pour décrire une propriété de type PT_CURRENCY. Pour plus d’informations sur les types de propriétés, voir [MAPI Property Type Overview](mapi-property-type-overview.md).
  
## <a name="see-also"></a>Voir aussi



[SPropValue](spropvalue.md)


[Structures MAPI](mapi-structures.md)

