---
title: PROP_ID
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.PROP_ID
api_type:
- COM
ms.assetid: 6ddaced5-49bb-41fe-95da-4e3300883bf7
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 228ea91969b35a1608dd6b3378b751312aa9c665
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33409131"
---
# <a name="prop_id"></a>PROP_ID

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Renvoie l’identificateur de propriété d’une balise de propriété spécifiée.
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapidefs.h  <br/> |
|Structure connexe :  <br/> |[SPropValue](spropvalue.md) <br/> |
   
```cpp
PROP_ID (ulPropTag)
```

## <a name="parameters"></a>Parameters

 _ulPropTag_
  
> Balise de propriété qui contient l’identificateur à retourner.
    
## <a name="remarks"></a>Remarques

Chaque balise de propriété contient le type de propriété dans le mot de bas ordre (bits 0 à 15) et l’identificateur de propriété dans le mot de haut niveau (bits 16 à 31). La **PROP_ID** macro extrait l’identificateur de propriété et le place en bits 0 à 15 de l’ensemble à retourner. Les autres bits de la valeur de retour sont des zéros. 
  
La **macro PROP_ID** peut être utilisée, par exemple, pour récupérer un identificateur à transmettre à [IMAPIProp::GetNamesFromIDs](imapiprop-getnamesfromids.md). **GetNamesFromIDs** récupère le nom de propriété associé à un identificateur pour une propriété nommée. 
  
## <a name="see-also"></a>Voir aussi



[SPropValue](spropvalue.md)


[Macros liées aux structures](macros-related-to-structures.md)

