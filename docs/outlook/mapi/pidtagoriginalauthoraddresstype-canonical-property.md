---
title: Propriété canonique PidTagOriginalAuthorAddressType
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.PidTagOriginalAuthorAddressType
api_type:
- COM
ms.assetid: 7cdedb1a-e441-469b-be50-2f18203eb30d
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 596e416624fb6f2bf1fdaef64c2179feb7787815
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33431413"
---
# <a name="pidtagoriginalauthoraddresstype-canonical-property"></a>Propriété canonique PidTagOriginalAuthorAddressType

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient le type d’adresse de l’auteur de la première version d’un message, c’est-à-dire le message avant d’être transmis ou de répondre.
  
|||
|:-----|:-----|
|Propriétés associées :  <br/> |PR_ORIGINAL_AUTHOR_ADDRTYPE, PR_ORIGINAL_AUTHOR_ADDRTYPE_A, PR_ORIGINAL_AUTHOR_ADDRTYPE_W  <br/> |
|Identificateur :  <br/> |0x0079  <br/> |
|Type de données :  <br/> |PT_STRING8, PT_UNICODE  <br/> |
|Domaine :  <br/> |Server  <br/> |
   
## <a name="remarks"></a>Remarques

Ces propriétés sont des exemples de propriétés d’adresse pour l’auteur d’un message. Lors de la première soumission du message, l’application cliente doit définir cette propriété sur la valeur de la propriété **PR_SENDER_ADDRTYPE** ([PidTagSenderAddressType](pidtagsenderaddresstype-canonical-property.md)). Elle n’est jamais modifiée lorsque le message est transmis ou à qui une réponse est répondue.
  
Les propriétés d’auteur d’origine permettent la conservation des informations en dehors du domaine de messagerie local. Lorsqu’un message arrive à partir d’un autre domaine de messagerie, tel qu’Internet, ces propriétés permettent de s’assurer que les informations d’origine ne sont pas perdues.
  
## <a name="related-resources"></a>Ressources connexes

### <a name="header-files"></a>Fichiers d’en-tête

Mapidefs.h
  
> Fournit des définitions de type de données.
    
Mapitags.h
  
> Contient les définitions des propriétés répertoriées en tant que propriétés associées.
    
## <a name="see-also"></a>Voir aussi



[Propriétés MAPI](mapi-properties.md)
  
[Propriétés canoniques MAPI](mapi-canonical-properties.md)
  
[Mappage des noms de propriétés canoniques aux noms MAPI](mapping-canonical-property-names-to-mapi-names.md)
  
[Mappage des noms MAPI aux noms de propriétés canoniques](mapping-mapi-names-to-canonical-property-names.md)

