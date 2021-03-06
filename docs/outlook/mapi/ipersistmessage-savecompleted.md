---
title: IPersistMessageSaveCompleted
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IPersistMessage.SaveCompleted
api_type:
- COM
ms.assetid: 83161011-90b4-49cb-9bcd-153a21a10977
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 021be209a2b2c891b668fa401500d6220619f9bd
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32317135"
---
# <a name="ipersistmessagesavecompleted"></a>IPersistMessage::SaveCompleted

**S’applique à** : Outlook 2013 | Outlook 2016 
  
Avertit le formulaire qu’une opération d’enregistrer a été effectuée. 
  
```cpp
HRESULT SaveCompleted(
  LPMESSAGE pMessage
);
```

## <a name="parameters"></a>Parameters

_pMessage_
  
> [in] Pointeur vers le message récemment enregistré.
    
## <a name="return-value"></a>Valeur renvoyée

S_OK 
  
> La notification a réussi.
    
E_INVALIDARG 
  
> Le _paramètre pMessage_ est NULL et le formulaire est à l’état [HandsOffFromNormal](handsofffromnormal-state.md) ou [HandsOffAfterSave.](handsoffaftersave-state.md) 
    
E_UNEXPECTED 
  
> Le formulaire n’est pas dans l’un des états suivants :
    
   - HandsOffFromNormal
    
   - HandsOffAfterSave
    
   - [NoScribble](noscribble-state.md)
    
## <a name="remarks"></a>Remarques

La **méthode IPersistMessage::SaveCompleted** est appelée par une visionneuse de formulaires pour informer le formulaire que toutes les modifications en attente ont été enregistrées. **SaveCompleted doit** être appelé uniquement lorsque le formulaire se trouve dans l’un des états suivants : 
  
- HandsOffFromNormal
    
- HandsOffAfterSave
    
- NoScribble
    
## <a name="notes-to-implementers"></a>Remarques pour les responsables de l’implémentation

Il existe plusieurs actions possibles que la méthode **SaveCompleted** peut effectuer, en fonction du contenu du paramètre de pointeur de message et de l’état dans le message. Toutefois, lorsqu’une action réussit, enregistrez toujours l’état actuel du message vers qui pointe le paramètre _pMessage_ et transition du formulaire vers son [état Normal.](normal-state.md) 
  
Le tableau suivant décrit les conditions qui affectent les actions que vous devez prendre dans votre implémentation de **SaveCompleted**.
  
|**Condition**|**Action**|
|:-----|:-----|
|Le  _paramètre pMessage_ est NULL et le paramètre  _fSameAsLoad_ de la méthode [IPersistMessage::Save](ipersistmessage-save.md) est définie sur TRUE.  <br/> |Appelez [la méthode IMAPIViewAdviseSink::OnSaved](imapiviewadvisesink-onsaved.md) de toutes les visionneuses inscrites, marquez le formulaire comme propre et renvoyez S_OK.  <br/> |
|Le  _paramètre pMessage_ est NULL et le paramètre  _fSameAsLoad_ de la méthode **IPersistMessage::Save** est définie sur FALSE.  <br/> |Elles retournent S_OK.  <br/> |
|Le formulaire est à l’état HandsOffFromNormal.  <br/> |Relâchez le message actuel et remplacez-le par le message pointé par _le paramètre pMessage._ Appelez la méthode [IUnknown::AddRef](https://msdn.microsoft.com/library/b4316efd-73d4-4995-b898-8025a316ba63%28Office.15%29.aspx) du message de remplacement et renvoyez S_OK.  <br/> |
|Le formulaire est à l’état HandsOffAfterSave.  <br/> |Appelez **la méthode IMAPIViewAdviseSink::OnSaved** de toutes les visionneuses inscrites, marquez le formulaire comme propre et renvoyez S_OK.  <br/> |
|Le formulaire est dans [l’état NoScribble.](noscribble-state.md)  <br/> |Relâchez le message actuel et remplacez-le par le message pointé  _par pMessage_. Appelez la méthode **IUnknown::AddRef** du message de remplacement. Appelez **la méthode IMAPIViewAdviseSink::OnSaved** de toutes les visionneuses inscrites, marquez le formulaire comme propre et renvoyez S_OK.  <br/> |
|Le formulaire est dans l’un des états HandsOff et le  _paramètre pMessage_ est définie sur NULL.  <br/> |Renvoyer E_INVALIDARG.  <br/> |
|Le formulaire est dans un état autre que l’un des états HandsOff ou NoScribble.  <br/> |Renvoyer E_UNEXPECTED.  <br/> |
   
Pour plus d’informations sur l’enregistrement des objets de stockage, voir la documentation des méthodes [IPersistStorage::SaveCompleted](https://docs.microsoft.com/windows/desktop/api/objidl/nf-objidl-ipersiststorage-savecompleted) ou [IPersistFile::SaveCompleted.](https://docs.microsoft.com/windows/desktop/api/objidl/nf-objidl-ipersistfile-savecompleted) 
  
## <a name="see-also"></a>Voir aussi

- [IPersistMessage : IUnknown](ipersistmessageiunknown.md)
- [États de formulaire](form-states.md)
