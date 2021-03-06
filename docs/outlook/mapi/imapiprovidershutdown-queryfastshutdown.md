---
title: IMAPIProviderShutdownQueryFastShutdown
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IMAPIProviderShutdown.QueryFastShutdown
api_type:
- COM
ms.assetid: 12069912-4b87-4945-9123-51106e0d2d54
description: 'Derniére modification : samedi 23 juillet 2011'
ms.openlocfilehash: 50b49761cf5923b11a450cbce7b7991f5ddd4d82
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33418217"
---
# <a name="imapiprovidershutdownqueryfastshutdown"></a>IMAPIProviderShutdown::QueryFastShutdown

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Interroge le fournisseur MAPI pour obtenir une prise en charge de l’arrêt rapide. 
  
```cpp
HRESULT QueryFastShutdown ();
```

## <a name="return-value"></a>Valeur renvoyée

S_OK
  
> Le fournisseur MAPI prend en charge le client MAPI pour un arrêt rapide.
    
MAPI_E_NO_SUPPORT
  
> Le fournisseur MAPI ne prend pas en charge le client MAPI pour un arrêt rapide.
    
## <a name="remarks"></a>Remarques

Les fournisseurs MAPI qui n’ont pas besoin de prendre en charge l’arrêt rapide du client doivent toujours implémenter l’interface [IMAPIProviderShutdown](imapiprovidershutdowniunknown.md) et faire en quoi la méthode **IMAPIProviderShutdown::QueryFastShutdown** doit être MAPI_E_NO_SUPPORT. Pour Outlook client MAPI, cela entraîne Outlook attendre que toutes les références externes soient libérées avant de se quitter. 
  
Selon le paramètre de Registre Windows de l’utilisateur pour un arrêt rapide, le fait de ne pas implémenter l’interface **IMAPIProviderShutdown** n’empêche pas nécessairement un arrêt rapide du client. 
  
## <a name="see-also"></a>Voir aussi



[IMAPIProviderShutdown : IUnknown](imapiprovidershutdowniunknown.md)


[Arrêt du client dans MAPI](client-shutdown-in-mapi.md)

