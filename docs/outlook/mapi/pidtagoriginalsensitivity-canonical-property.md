---
title: Propriété canonique PidTagOriginalSensitivity
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.PidTagOriginalSensitivity
api_type:
- COM
ms.assetid: 70a87cf8-2011-4669-90fd-2711c3352e30
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: e712a4cc49541ee4330f479d7a03af323bdbc887
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32341229"
---
# <a name="pidtagoriginalsensitivity-canonical-property"></a>Propriété canonique PidTagOriginalSensitivity

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient la valeur de sensibilité attribuée par l’expéditeur de la première version d’un message, c’est-à-dire, le message avant d’être transmis ou de répondre.
  
|||
|:-----|:-----|
|Propriétés associées :  <br/> |PR_ORIGINAL_SENSITIVITY  <br/> |
|Identificateur :  <br/> |0x002E  <br/> |
|Type de données :  <br/> |PT_LONG  <br/> |
|Domaine :  <br/> |Messagerie générale  <br/> |
   
## <a name="remarks"></a>Remarques

Une application cliente doit définir cette propriété sur la même valeur que la propriété **PR_SENSITIVITY** ([PidTagSensitivity](pidtagsensitivity-canonical-property.md)) lorsque le message est envoyé pour la première fois. Elle ne doit jamais être modifiée par la suite.
  
Cette propriété est utilisée par le fournisseur de transport pour protéger la sensibilité sur les entrées copiées. Elle permet, par exemple, de bloquer la modification du texte du message d’origine dans un avant ou d’une réponse à un message qui a été initialement marqué **comme SENSITIVITY_PRIVATE**.
  
## <a name="related-resources"></a>Ressources connexes

### <a name="protocol-specifications"></a>Spécifications de protocole

[[MS-OXPROPS]](https://msdn.microsoft.com/library/f6ab1613-aefe-447d-a49c-18217230b148%28Office.15%29.aspx)
  
> Fournit des références aux spécifications Exchange Server protocole.
    
[[MS-OXOMSG]](https://msdn.microsoft.com/library/daa9120f-f325-4afb-a738-28f91049ab3c%28Office.15%29.aspx)
  
> Spécifie les propriétés et opérations autorisées sur les objets de message électronique.
    
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

