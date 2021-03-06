---
title: Propriété canonique PidTagJunkThreshold
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- PidTagJunkThreshold
api_type:
- HeaderDef
ms.assetid: 8067e2b5-02df-4b96-8f66-509f5a48c8aa
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 078dfcc7c24870cf95a2a4b2385c34fbeb64fac0
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32326851"
---
# <a name="pidtagjunkthreshold-canonical-property"></a>Propriété canonique PidTagJunkThreshold

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Indique la façon dont le courrier entrant doit être envoyé de manière agressive au dossier Courrier indésirable.
  
|||
|:-----|:-----|
|Propriétés associées :  <br/> |PR_JUNK_THRESHOLD  <br/> |
|Identificateur :  <br/> |0x6101  <br/> |
|Type de données :  <br/> |PT_LONG  <br/> |
|Domaine :  <br/> |Courrier indésirable  <br/> |
   
## <a name="remarks"></a>Remarques

Cette propriété correspond au paramètre de filtre élevé/faible/aucun. La valeur « 0xFFFFFFFF » indique que le filtrage du courrier indésirable ne doit pas être appliqué, mais que les listes d’adresses bloqués doivent toujours être appliquées. La valeur « 0x80000000 » indique que tous les messages sont du courrier indésirable à l’exception des messages provenant d’expéditeurs de la liste des expéditeurs fiables ou envoyés à des destinataires de la liste des destinataires fiables. Les valeurs sont les suivantes :
  
|**Valeur**|**Description**|
|:-----|:-----|
|0xFFFFFFFF  <br/> |Aucun filtrage du courrier indésirable  <br/> |
|0x00000006  <br/> |Filtrage du courrier indésirable faible  <br/> |
|0x00000003  <br/> |Filtrage du courrier indésirable élevé  <br/> |
|0x80000000  <br/> |Listes de confiance uniquement  <br/> |
   
## <a name="related-resources"></a>Ressources connexes

### <a name="protocol-specifications"></a>Spécifications de protocole

[[MS-OXPROPS]](https://msdn.microsoft.com/library/f6ab1613-aefe-447d-a49c-18217230b148%28Office.15%29.aspx)
  
> Fournit des références aux spécifications Exchange Server protocole.
    
[[MS-OXCSPAM]](https://msdn.microsoft.com/library/522f8587-4aed-4cd6-831b-40bd87862189%28Office.15%29.aspx)
  
> Permet la gestion des listes d’adresses de courriers électroniques indésirables et la détermination des listes d’adresses de courrier indésirable.
    
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

