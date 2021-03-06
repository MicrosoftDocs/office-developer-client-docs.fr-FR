---
title: CloseIMsgSession
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.CloseIMsgSession
api_type:
- COM
ms.assetid: a0a17309-fc59-4822-be9b-b6f623b68bb1
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 877bebf0a156c99907505d815ca8d36a4b398678
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33412036"
---
# <a name="closeimsgsession"></a>CloseIMsgSession

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Ferme une session de message et tous les messages créés au sein de cette session. 
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Imessage.h  <br/> |
|Implémenté par :  <br/> |MAPI  <br/> |
|Appelé par :  <br/> |Applications clientes et fournisseurs de services  <br/> |
   
```cpp
VOID CloseIMsgSession(
  LPMSGSESS lpMsgSess
);
```

## <a name="parameters"></a>Parameters

 _lpMsgSess_
  
> [in] Pointeur vers l’objet de session de message obtenu à l’aide de la fonction [OpenIMsgSession](openimsgsession.md) au début de la session de message. 
    
## <a name="return-value"></a>Valeur renvoyée

Aucun.
  
## <a name="remarks"></a>Remarques

Une session de message est utilisée par les applications clientes et les fournisseurs de services qui souhaitent traiter plusieurs objets **IMessage** MAPI associés créés au-dessus des objets OLE **IStorage** sous-jacents. Le client ou le fournisseur utilise les fonctions [OpenIMsgSession](openimsgsession.md) et **CloseIMsgSession** pour encapsuler la création de tels messages dans une session de message. Une fois la session de message ouverte, le client ou le fournisseur lui transmet un pointeur dans un appel à [OpenIMsgOnIStg](openimsgonistg.md) pour créer un objet **IMessage**-on- **IStorage.** 
  
Une session de message assure le suivi de tous les objets **IMessage**-on- **IStorage** ouverts pendant la durée de la session, en plus de toutes les pièces jointes et autres propriétés des messages. Lorsqu’un client ou un fournisseur appelle **CloseIMsgSession,** il ferme tous ces objets. L’appel de **CloseIMsgSession** est le seul moyen de fermer des objets **IMessage**-on- **IStorage.** 
  

