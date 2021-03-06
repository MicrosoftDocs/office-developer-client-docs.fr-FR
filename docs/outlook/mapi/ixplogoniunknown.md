---
title: IXPLogon IUnknown
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IXPLogon
api_type:
- COM
ms.assetid: 4d24ecaf-11d0-4362-8207-be3407736d7b
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 46f4e3fc8f554f332ab9b1d8a6cb33e9e21dd9a5
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33432533"
---
# <a name="ixplogon--iunknown"></a>IXPLogon : IUnknown

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Donne aupooler MAPI l’accès à un fournisseur de transport. 
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapispi.h  <br/> |
|Exposé par :  <br/> |Objets d’logo de transport  <br/> |
|Implémenté par :  <br/> |Fournisseurs de transport  <br/> |
|Appelé par :  <br/> |Lepooler MAPI  <br/> |
|Identificateur d’interface :  <br/> |IID_IXPLogon  <br/> |
|Type de pointeur :  <br/> |LXPLOGON  <br/> |
   
## <a name="vtable-order"></a>Ordre des vtables

|||
|:-----|:-----|
|[AddressTypes](ixplogon-addresstypes.md) <br/> |Renvoie les types de destinataires gérés par le fournisseur de transport.  <br/> |
|**RegisterOptions** <br/> | *Non pris en charge ou documenté.*  <br/> |
|[TransportNotify](ixplogon-transportnotify.md) <br/> |Signale l’occurrence d’un événement sur lequel le fournisseur de transport a demandé une notification.  <br/> |
|[Inactif](ixplogon-idle.md) <br/> |Indique que le système est inactif, ce qui permet au fournisseur de transport d’effectuer des opérations de faible priorité.  <br/> |
|[TransportLogoff](ixplogon-transportlogoff.md) <br/> |Lance le processus de ffage de logo.  <br/> |
|[SubmitMessage](ixplogon-submitmessage.md) <br/> |Indique que lepooler MAPI a un message à remettre au fournisseur de transport.  <br/> |
|[EndMessage](ixplogon-endmessage.md) <br/> |Informe le fournisseur de transport que lepooler MAPI a terminé son traitement sur un message sortant.  <br/> |
|[Sondage](ixplogon-poll.md) <br/> |Indique si le fournisseur de transport a reçu un ou plusieurs messages entrants.  <br/> |
|[StartMessage](ixplogon-startmessage.md) <br/> |Lance le transfert d’un message entrant du fournisseur de transport vers lepooler MAPI.  <br/> |
|[OpenStatusEntry](ixplogon-openstatusentry.md) <br/> |Ouvre l’objet d’état du fournisseur de transport.  <br/> |
|[ValidateState](ixplogon-validatestate.md) <br/> |Vérifie l’état externe du fournisseur de transport.  <br/> |
|[FlushQueues](ixplogon-flushqueues.md) <br/> |Demande au fournisseur de transport de remettre immédiatement tous les messages entrants ou sortants en attente.  <br/> |
   
## <a name="see-also"></a>Voir aussi



[Interfaces MAPI](mapi-interfaces.md)

