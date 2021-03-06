---
title: Structure des fournisseurs de banques de messages
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
localization_priority: Normal
api_type:
- COM
ms.assetid: 064b2fc1-e690-43e6-95d3-a61438115de5
description: 'Derniére modification : samedi 23 juillet 2011'
ms.openlocfilehash: eda62a4cd31e0de695d52391a6717e7a0f5ea581
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33426421"
---
# <a name="structure-of-message-store-providers"></a>Structure des fournisseurs de banques de messages
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Un fournisseur de magasin de messages, lorsqu’il est en cours d’exécution en mémoire, est une interface [IMSProvider : IUnknown.](imsprovideriunknown.md) **L’interface IMSProvider** permet aux applications clientes et aupooler MAPI de se connecter et de se déconnecter de la boutique de messages. Les interfaces que les applications clientes et lepooler MAPI utilisent pour accéder aux dossiers et aux messages dans la boutique de messages sont les interfaces [IMSLogon](imslogoniunknown.md) et [IMsgStore.](imsgstoreimapiprop.md) Ces interfaces sont généralement créées lors de la première ouverture de session de la magasin de messages, bien que le point d’entrée [MSProviderInit](msproviderinit.md) de la DLL de la magasin de messages puisse également les créer. 
  
Étant donné que les interfaces **IMSLogon** et **IMsgStore** partagent certaines méthodes, il peut être plus facile de créer un objet de classe qui hérite de ces deux interfaces. Vous pouvez également implémenter ces interfaces dans des objets distincts et écrire des fonctions d’aide internes à votre DLL qui implémentent les méthodes partagées qui peuvent ensuite être appelées à partir des méthodes des interfaces **IMSLogon** et **IMsgStore.** 
  
L’illustration suivante montre un plan de haut niveau de la hiérarchie d’objets dans une magasin de messages en cours d’exécution.
  
**Hiérarchie d’objets de banque de messages**
  
![Hiérarchie d’objets de magasin de messages Hiérarchie](media/storeobj.gif "d’objets de magasin de messages")
  
## <a name="see-also"></a>Voir aussi

- [Développement d’un fournisseur de banque de messages MAPI](developing-a-mapi-message-store-provider.md)

