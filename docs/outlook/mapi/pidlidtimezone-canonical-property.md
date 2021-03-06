---
title: Propriété canonique PidLidTimeZone
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- PidLidTimeZone
api_type:
- COM
ms.assetid: ffbab371-1a1d-4aa4-ad31-17549a74513c
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: b62779567a7dbd298fdd313e90b13fb223e4e47e
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32319270"
---
# <a name="pidlidtimezone-canonical-property"></a>Propriété canonique PidLidTimeZone

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Spécifie des informations sur le fuseau horaire d’une réunion périodique.
  
|||
|:-----|:-----|
|Propriétés associées :  <br/> |LID_TIME_ZONE  <br/> |
|Jeu de propriétés :  <br/> |PSETID_Meeting  <br/> |
|ID long (LID) :  <br/> |0x0000000C  <br/> |
|Type de données :  <br/> |PT_LONG  <br/> |
|Domaine :  <br/> |Réunions  <br/> |
   
## <a name="remarks"></a>Remarques

Cette propriété est lue uniquement si la propriété **dispidApptRecur** ([PidLidAppointmentRecur](pidlidappointmentrecur-canonical-property.md)) n’est pas définie, mais si la propriété **LID_IS_RECURRING** ([PidLidIsRecurring](pidlidisrecurring-canonical-property.md)) est TRUE et la propriété **LID_IS_EXCEPTION** ([PidLidIsException](pidlidisexception-canonical-property.md)) est FALSE. Word inférieur spécifie un index dans un tableau qui contient des informations de fuseau horaire. Dans la partie supérieure de WORD, seul le bit le plus élevé est lu. Si ce bit est définie, le fuseau horaire référencé n’observera pas l’heure d’été , sinon les dates d’été détaillées dans [[MS-OXOCAL]](https://msdn.microsoft.com/library/09861fde-c8e4-4028-9346-e7c214cfdba1%28Office.15%29.aspx) seront suivies. 
  
## <a name="related-resources"></a>Ressources connexes

### <a name="protocol-specifications"></a>Spécifications de protocole

[[MS-OXPROPS]](https://msdn.microsoft.com/library/f6ab1613-aefe-447d-a49c-18217230b148%28Office.15%29.aspx)
  
> Fournit des définitions de jeu de propriétés et des références aux spécifications Exchange Server protocole.
    
[[MS-OXOCAL]](https://msdn.microsoft.com/library/09861fde-c8e4-4028-9346-e7c214cfdba1%28Office.15%29.aspx)
  
> Spécifie les propriétés et les opérations pour les messages de rendez-vous, de demande de réunion et de réponse.
    
### <a name="header-files"></a>Fichiers d’en-tête

Mapidefs.h
  
> Fournit des définitions de type de données.
    
## <a name="see-also"></a>Voir aussi



[Propriétés MAPI](mapi-properties.md)
  
[Propriétés canoniques MAPI](mapi-canonical-properties.md)
  
[Mappage des noms de propriétés canoniques aux noms MAPI](mapping-canonical-property-names-to-mapi-names.md)
  
[Mappage des noms MAPI aux noms de propriétés canoniques](mapping-mapi-names-to-canonical-property-names.md)

