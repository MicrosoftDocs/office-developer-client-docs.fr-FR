---
title: IID
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.IID
api_type:
- COM
ms.assetid: fa5498ab-2f8a-42f8-ba9d-1d555768594f
description: 'Derniére modification : samedi 23 juillet 2011'
ms.openlocfilehash: 5605de7dbcc18197748713bcf909839690d7259f
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33411595"
---
# <a name="iid"></a>IID

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Décrit une structure [GUID](guid.md) utilisée pour décrire un identificateur pour une interface MAPI. 
  
```cpp
typedef struct _GUID
{
  unsigned long Data1;
  unsigned short Data2;
  unsigned short Data3;
  unsigned char Data4[8];
} GUID;

```

## <a name="members"></a>Members

Voir la structure **GUID.** 
  
## <a name="remarks"></a>Remarques

Une **structure IID** permet d’identifier de manière unique une interface MAPI et d’associer une interface particulière à un objet. Par exemple, lorsqu’un client appelle [IMAPISession::OpenEntry](imapisession-openentry.md) pour ouvrir un dossier, le client définit le paramètre _lpInterface_ pour qu’il pointe vers un **IID** représentant l’interface [IMAPIFolder.](imapifolderimapicontainer.md) MAPI définit **IMAPIFolderIID** à IID_IMAPIFolder. Les structures **IID** sont également utilisées pour identifier de manière unique les interfaces OLE. 
  
Toutes les structures **IID** spécifiques pour les interfaces MAPI sont définies dans le fichier d’en-tête Mapiguid.h. 
  
## <a name="see-also"></a>Voir aussi



[GUID](guid.md)


[Structures MAPI](mapi-structures.md)

