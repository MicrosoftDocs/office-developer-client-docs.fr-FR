---
title: Propriété canonique PidTagRtfSyncPrefixCount
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.PidTagRtfSyncPrefixCount
api_type:
- COM
ms.assetid: c2b15ac5-9e89-4ee2-812d-102d0b2ac56e
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 61683534504b7451f126591af149d11306cff1bd
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32357861"
---
# <a name="pidtagrtfsyncprefixcount-canonical-property"></a>Propriété canonique PidTagRtfSyncPrefixCount

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient le nombre de caractères ignorés qui apparaissent avant les caractères significatifs du message.
  
|||
|:-----|:-----|
|Propriétés associées :  <br/> |PR_RTF_SYNC_PREFIX_COUNT  <br/> |
|Identificateur :  <br/> |0x1010  <br/> |
|Type de données :  <br/> |PT_LONG  <br/> |
|Domaine :  <br/> |Message MAPI  <br/> |
   
## <a name="remarks"></a>Remarques

Le nombre de caractères de préfixe n’inclut pas d’espace blanc.
  
Cette propriété est une propriété auxiliaire rtf (Rich Text Format). Les propriétés auxiliaires RTF sont utilisées par la [fonction RTFSync](rtfsync.md) et ne sont pas destinées à être utilisées directement par les applications clientes. 
  
## <a name="related-resources"></a>Ressources connexes

### <a name="protocol-specifications"></a>Spécifications de protocole

[[MS-OXPROPS]](https://msdn.microsoft.com/library/f6ab1613-aefe-447d-a49c-18217230b148%28Office.15%29.aspx)
  
> Fournit des références aux spécifications Exchange Server protocole.
    
[[MS-OXTNEF]](https://msdn.microsoft.com/library/1f0544d7-30b7-4194-b58f-adc82f3763bb%28Office.15%29.aspx)
  
> Encode et décode les objets de message et de pièce jointe dans une représentation de flux efficace.
    
### <a name="header-files"></a>Fichiers d’en-tête

Mapidefs.h
  
> Fournit des définitions de type de données.
    
Mapitags.h
  
> Contient les définitions des propriétés répertoriées en tant que noms de remplacement.
    
## <a name="see-also"></a>Voir aussi



[Propriétés MAPI](mapi-properties.md)
  
[Propriétés canoniques MAPI](mapi-canonical-properties.md)
  
[Mappage des noms de propriétés canoniques aux noms MAPI](mapping-canonical-property-names-to-mapi-names.md)
  
[Mappage des noms MAPI aux noms de propriétés canoniques](mapping-mapi-names-to-canonical-property-names.md)

