---
title: IProxyStoreObjectUnwrapNoRef
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IProxyStoreObject.UnwrapNoRef
api_type:
- COM
ms.assetid: 1122b6e0-e7e1-e68a-e090-435777343d04
description: 'Derniére modification : samedi 23 juillet 2011'
ms.openlocfilehash: ef9f506c1a95fec86c7f092b0299198e6149d3ba
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32320152"
---
# <a name="iproxystoreobjectunwrapnoref"></a>IProxyStoreObject::UnwrapNoRef

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Obtient un pointeur vers un objet de magasin IMAP (Internet Message Access Protocol) non veloppé qui permet d’accéder au fichier PST (Personal Folders) sous-jacent sans avoir à demander la synchronisation et à télécharger les éléments.
  
```cpp
HRESULT IProxyStoreObject::UnwrapNoRef (     LPVOID *ppvObject ); 
```

## <a name="parameters"></a>Parameters

 _ppvObject_
  
> [out] Pointeur vers un objet de magasin [IMsgStore : IMAPIProp](imsgstoreimapiprop.md) qui est déballé. 
    
## <a name="return-values"></a>Valeurs de retour

S_OK
  
- L’appel a réussi et un pointeur vers une interface non enveloppée a été renvoyé dans  _ppvObject_.
    
## <a name="remarks"></a>Remarques

Sans avoir d’abord désécrit un magasin IMAP, l’accès à un message dans la boutique peut forcer une synchronisation qui tente de télécharger l’intégralité du message. L’utilisation de la store non enveloppée permet d’accéder au message dans son état actuel sans déclencher de téléchargement.
  
Étant donné que **UnwrapNoRef** n’incrémente pas le nombre de références pour ce nouveau pointeur vers l’objet store non enveloppé, après avoir correctement appelé **UnwrapNoRef,** vous devez appeler [IUnknown::AddRef](https://msdn.microsoft.com/library/ms691379%28v=VS.85%29.aspx) pour conserver le nombre de références. 
  
## <a name="see-also"></a>Voir aussi



[IProxyStoreObject](iproxystoreobject.md)

