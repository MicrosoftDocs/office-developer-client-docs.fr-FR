---
title: SLongArray
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.SLongArray
api_type:
- COM
ms.assetid: 57435634-202d-4998-9931-4562f1a66f5f
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 9b1c5a09a60240efa9d4fa117f0d8fe8113169d5
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33414528"
---
# <a name="slongarray"></a>SLongArray

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient un tableau de types de valeur LONG utilisés pour décrire une propriété de type PT_MV_LONG. 
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapidefs.h  <br/> |
   
```cpp
typedef struct _SLongArray
{
  ULONG cValues;
  LONG FAR *lpl;
} SLongArray;

```

## <a name="members"></a>Members

 **cValues**
  
> Nombre de valeurs dans le tableau pointées par le **membre lpl.** 
    
 **lpl**
  
> Pointeur vers un tableau de valeurs LONG.
    
## <a name="remarks"></a>Remarques

Pour plus d’informations sur PT_MV_LONG, voir [Liste des types de propriétés.](property-types.md)
  
## <a name="see-also"></a>Voir aussi



[SPropValue](spropvalue.md)


[Structures MAPI](mapi-structures.md)

