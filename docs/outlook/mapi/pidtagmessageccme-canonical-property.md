---
title: Propriété canonique PidTagMessageCcMe
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- PidTagMessageCcMe
api_type:
- HeaderDef
ms.assetid: 7310a0f2-a109-40a4-99bf-e963d754a067
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 7605739dd6d0f0205a1a4f09eb8c45d235c0c179
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32329735"
---
# <a name="pidtagmessageccme-canonical-property"></a>Propriété canonique PidTagMessageCcMe

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient TRUE si cet utilisateur de messagerie est spécifiquement nommé en tant que destinataire de copie carbone (CC) de ce message et ne fait pas partie d’une liste de distribution. 
  
|||
|:-----|:-----|
|Propriétés associées :  <br/> |PR_MESSAGE_CC_ME  <br/> |
|Identificateur :  <br/> |0x0058  <br/> |
|Type de données :  <br/> |PT_BOOLEAN  <br/> |
|Domaine :  <br/> |Messagerie générale  <br/> |
   
## <a name="remarks"></a>Remarques

Cette propriété offre un moyen pratique de déterminer si le nom d’utilisateur apparaît explicitement dans la liste des destinataires en copie carbone, sans examiner toutes les entrées de la liste. 
  
Cette propriété permet également la gestion automatisée des messages reçus au moment de la réception. À l’option du fournisseur de transport, cette propriété contient false ou n’est pas définie si l’utilisateur de messagerie n’est pas répertorié directement dans la table des destinataires. 
  
La remise des messages résultant de l’extension de la liste de distribution ou d’une désignation de copie carbone non voyante n’entraîne pas la mise en place de cette propriété. Le destinataire doit être nommé explicitement. 
  
Les messages non envoyés ne définissent généralement pas cette propriété, **PR_MESSAGE_RECIP_ME** ([PidTagMessageRecipientMe](pidtagmessagerecipientme-canonical-property.md)) ou **PR_MESSAGE_TO_ME** ([PidTagMessageToMe](pidtagmessagetome-canonical-property.md)). S’ils sont présents dans des messages accessibles par l’utilisateur dans des magasins de messages publics, dans des magasins privés d’autres utilisateurs, dans des fichiers sur disque ou incorporés à d’autres messages reçus, ils contiennent généralement les valeurs pour lesquelles ils ont été définies la dernière fois qu’un fournisseur de transport a remis le message. 
  
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

