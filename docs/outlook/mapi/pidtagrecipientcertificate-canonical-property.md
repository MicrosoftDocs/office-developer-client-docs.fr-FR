---
title: Propriété canonique PidTagRecipientCertificate
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.PidTagRecipientCertificate
api_type:
- COM
ms.assetid: 7c5c749e-5463-4935-85b5-32219d06f782
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: c659b77767fddc4c783732082c2eb65c68af8dbf
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33431665"
---
# <a name="pidtagrecipientcertificate-canonical-property"></a>Propriété canonique PidTagRecipientCertificate

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient le certificat ASN.1 d’un destinataire de message à utiliser dans un état.
  
|||
|:-----|:-----|
|Propriétés associées :  <br/> |PR_RECIPIENT_CERTIFICATE  <br/> |
|Identificateur :  <br/> |0x0C13  <br/> |
|Type de données :  <br/> |PT_BINARY  <br/> |
|Domaine :  <br/> |Destinataire MAPI  <br/> |
   
## <a name="remarks"></a>Remarques

Cette propriété est une copie de la propriété PR_USER_CERTIFICATE **(** [PidTagUserCertificate](pidtagusercertificate-canonical-property.md)) du destinataire à utiliser dans un état. Elle peut être utilisée pour prouver à l’auteur que le destinataire a réellement reçu le message, ce qu’une note de remise n’indique pas nécessairement.
  
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

