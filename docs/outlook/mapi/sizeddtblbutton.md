---
title: SizedDtblButton
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.SizedDtblButton
api_type:
- COM
ms.assetid: ee73ced9-14d8-4a30-9b9f-d54ed9c3a454
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: a7d62d29638ae234667eb33a8103fb3a716afc32
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33421913"
---
# <a name="sizeddtblbutton"></a>SizedDtblButton

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Crée une structure nommée qui inclut une structure [DTBLBUTTON](dtblbutton.md) pour décrire un bouton et une étiquette d’une longueur spécifiée. 
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapidefs.h  <br/> |
|Structure connexe :  <br/> |**DTBLBUTTON** <br/> |
   
```cpp
SizedDtblButton (n, u)
```

## <a name="parameters"></a>Parameters

 _n_
  
> Longueur de l’étiquette à inclure dans la nouvelle structure.
    
 _u_
  
> Nom de la nouvelle structure.
    
## <a name="remarks"></a>Remarques

La nouvelle structure est créée avec les membres suivants :
  
```
DTBLBUTTON dtblbutton;
TCHAR lpszLabel[n];

```

## <a name="see-also"></a>Voir aussi



[DTBLBUTTON](dtblbutton.md)


[Macros liées aux structures](macros-related-to-structures.md)

