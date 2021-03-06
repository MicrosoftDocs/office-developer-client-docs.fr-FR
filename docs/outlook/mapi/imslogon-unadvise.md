---
title: IMSLogonUnadvise
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IMSLogon.Unadvise
api_type:
- COM
ms.assetid: 440d61c4-b69a-4010-a22b-0c9c5c376fbc
description: 'Derniére modification : samedi 23 juillet 2011'
ms.openlocfilehash: 9172d4956e78ac31cd15d69e70d05c127a474ca5
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32317275"
---
# <a name="imslogonunadvise"></a>IMSLogon::Unadvise

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Supprime l’inscription d’un objet pour la notification des modifications de magasin de messages précédemment établies à l’aide d’un appel à la méthode [IMSLogon::Advise.](imslogon-advise.md) 
  
```cpp
HRESULT Unadvise(
  ULONG ulConnection
);
```

## <a name="parameters"></a>Parameters

 _ulConnection_
  
> [in] Numéro de la connexion d’inscription renvoyée par un appel **à IMSLogon::Advise**.
    
## <a name="return-value"></a>Valeur renvoyée

S_OK 
  
> L'appel a r�ussi et a renvoy� la valeur attendue ou les valeurs.
    
## <a name="remarks"></a>Remarques

Les fournisseurs de magasins de messages implémentent la méthode **IMSLogon::Unadvise** pour libérer le pointeur vers l’objet de réception de notification transmis dans le paramètre  _lpAdviseSink_ dans l’appel précédent à **IMSLogon::Advise,** ce qui annule l’inscription d’une notification. Dans le cadre de l’ignorer, la méthode [IUnknown::Release](https://msdn.microsoft.com/library/ms682317%28v=VS.85%29.aspx) de l’objet est appelée. En règle générale, **Release** est appelé pendant **l’appel Unadvise.** Toutefois, si un autre thread est en train d’appeler la méthode [IMAPIAdviseSink::OnNotify](imapiadvisesink-onnotify.md) pour l’objet de sink de conseil, l’appel de publication est retardé jusqu’à ce que la méthode **OnNotify** renvoie.  
  
## <a name="see-also"></a>Voir aussi



[IMAPIAdviseSink::OnNotify](imapiadvisesink-onnotify.md)
  
[IMSLogon::Advise](imslogon-advise.md)
  
[IMSLogon : IUnknown](imslogoniunknown.md)

