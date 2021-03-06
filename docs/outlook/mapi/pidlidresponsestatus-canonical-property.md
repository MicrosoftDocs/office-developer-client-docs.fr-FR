---
title: Propri t canonique PidLidResponseStatus
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- PidLidResponseStatus
api_type:
- COM
ms.assetid: e56142fd-204b-497e-83b9-59f9acda6cb4
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 8c8e284752c6fd47eb7a8f227e2dbfeceebea596
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32345758"
---
# <a name="pidlidresponsestatus-canonical-property"></a>Propri t canonique PidLidResponseStatus

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Spécifie l’état de réponse d’un participant.
  
|||
|:-----|:-----|
|Propriétés associées :  <br/> |dispidResponseStatus  <br/> |
|Jeu de propriétés :  <br/> |PSETID_Appointment  <br/> |
|ID long (LID) :  <br/> |0x00008218  <br/> |
|Type de données :  <br/> |PT_LONG  <br/> |
|Domaine :  <br/> |Réunions  <br/> |
   
## <a name="remarks"></a>Remarques

L’état de la réponse doit être l’une des valeurs du tableau ci-dessous.
  
|**État de la réponse**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|respNone  <br/> |0x00000000  <br/> |Aucune réponse n’est requise pour cet objet. C’est le cas pour les objets de rendez-vous et les objets de réponse de réunion.  <br/> |
|respOrganized  <br/> |0x00000001  <br/> |Cette réunion appartient à l’organisateur.  <br/> |
|respTentative  <br/> |0x00000002  <br/> |Cette valeur sur la réunion du participant indique que le participant a provisoirement accepté la demande de réunion.  <br/> |
|respAccepted  <br/> |0x00000003  <br/> |Cette valeur sur la réunion du participant indique que le participant a accepté la demande de réunion.  <br/> |
|respDeclined  <br/> |0x00000004  <br/> |Cette valeur sur la réunion du participant indique que le participant a refusé la demande de réunion.  <br/> |
|respNotResponded  <br/> |0x00000005  <br/> |Cette valeur sur la réunion du participant indique que le participant n’a pas encore répondu. Cette valeur est sur la demande de réunion, la mise à jour de réunion et l’annulation de la réunion.  <br/> |
   
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

