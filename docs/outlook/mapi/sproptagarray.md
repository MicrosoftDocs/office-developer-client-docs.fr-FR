---
title: SPropTagArray
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.SPropTagArray
api_type:
- COM
ms.assetid: 4a9e1579-bebe-4a51-8ced-6dba9c3bcb63
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 9a5be98298ab1f9333ac1c223a6ef594e60dd86a
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33430699"
---
# <a name="sproptagarray"></a>SPropTagArray

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient un tableau de balises de propriétés. 
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapidefs.h  <br/> |
|Macros associées :  <br/> |[CbNewSPropTagArray](cbnewsproptagarray.md), [CbSPropTagArray](cbsproptagarray.md), [SizedSPropTagArray](sizedsproptagarray.md) <br/> |
   
```cpp
typedef struct _SPropTagArray
{
  ULONG cValues;
  ULONG aulPropTag[MAPI_DIM];
} SPropTagArray, FAR *LPSPropTagArray;

```

## <a name="members"></a>Members

 **cValues**
  
> Nombre de balises de propriété dans le tableau indiqué par le **membre aulPropTag.** 
    
 **aulPropTag**
  
> Tableau de balises de propriété.
    
## <a name="remarks"></a>Remarques

Une balise de propriété est un nombre d’éléments non signés 32 bits composé de deux parties : 
  
- Identificateur de l’ordre élevé de 16 bits.
    
- Type de 16 bits de bas ordre.
    
L’identificateur est une valeur numérique dans une plage particulière. MAPI définit des plages pour les identificateurs afin de décrire l’utilisation de la propriété et la personne responsable de sa maintenance. MAPI définit des contraintes pour chacune des balises de propriété qu’il prend en charge dans le fichier d’en-tête Mapitags.h.
  
Le type indique le format de la valeur de la propriété. MAPI définit des constantes pour chacun des types de propriétés qu’il prend en charge dans le fichier d’en-tête Mapidefs.h. 
  
Pour plus d’informations sur les balises de propriété et leurs composants, consultez l’une des rubriques suivantes : 
  
[Balises de propriété MAPI](mapi-property-tags.md)
  
[Vue d’ensemble de l’identificateur de propriété MAPI](mapi-property-identifier-overview.md)
  
[Vue d’ensemble du type de propriété MAPI](mapi-property-type-overview.md)
  
Pour obtenir la liste complète des types de propriétés à valeur unique et à valeurs multiples, voir l’annexe, Identificateurs de [propriétés et Types.](property-identifiers-and-types.md) 
  
## <a name="see-also"></a>Voir aussi



[Structures MAPI](mapi-structures.md)

