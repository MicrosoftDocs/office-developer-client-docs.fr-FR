---
title: FLAGLIST
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.FLAGLIST
api_type:
- COM
ms.assetid: b4c0655c-1a3a-4f89-a977-0431db596512
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: a5e508f5f7e6554a115517da87a8eac39f39aecf
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33412974"
---
# <a name="flaglist"></a>FLAGLIST

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient une liste d’indicateurs utilisés pour indiquer l’état des entrées d’adresse pendant le processus de résolution de noms.
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapidefs.h  <br/> |
   
```cpp
typedef struct
{
  ULONG cFlags;
  ULONG ulFlags[MAPI_DIM];
} FlagList, FAR * LPFlagList;

```

## <a name="members"></a>Members

 **cFlags**
  
> Nombre d’indicateurs définis par MAPI dans la liste.
    
 **ulFlags**
  
> Tableau d’indicateurs qui fournit l’état de l’opération de résolution de noms pour un destinataire. Les indicateurs suivants peuvent être définies :
    
MAPI_AMBIGUOUS 
  
> Le destinataire a été résolu, mais pas en un identificateur d’entrée unique. Les autres conteneurs de carnet d’adresses ne doivent pas essayer de résoudre ce destinataire. 
    
MAPI_RESOLVED 
  
> Le destinataire a été résolu en un identificateur d’entrée unique. Les autres conteneurs de carnet d’adresses ne doivent pas essayer de résoudre ce destinataire. 
    
MAPI_UNRESOLVED 
  
> L’entrée n’a pas été résolue. Les autres conteneurs de carnet d’adresses doivent essayer de résoudre ce destinataire.
    
## <a name="remarks"></a>Remarques

La structure **FLAGLIST** est utilisée comme paramètre pour [IABContainer::ResolveNames](iabcontainer-resolvenames.md). Chacun des destinataires à résoudre est inclus dans une structure [ADRLIST.](adrlist.md) Lorsque le conteneur de carnet d’adresses tente de résoudre chaque destinataire, il définit l’indicateur approprié dans l’entrée correspondante dans la structure **FLAGLIST.** Toutes les entrées de la structure **FLAGLIST** sont dans le même ordre que les entrées de la structure **ADRLIST.** Cela facilite l’association d’un paramètre d’indicateur à un destinataire. 
  
## <a name="see-also"></a>Voir aussi



[ADRLIST](adrlist.md)
  
[IABContainer::ResolveNames](iabcontainer-resolvenames.md)


[Structures MAPI](mapi-structures.md)

