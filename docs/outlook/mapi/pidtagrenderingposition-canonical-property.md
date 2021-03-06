---
title: Propriété canonique PidTagRenderingPosition
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.PidTagRenderingPosition
api_type:
- COM
ms.assetid: bce46687-17dc-4a3f-96be-303d8755158e
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: d463be4a14ecf478bdcbddc50b4ad9360829befc
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32355152"
---
# <a name="pidtagrenderingposition-canonical-property"></a>Propriété canonique PidTagRenderingPosition

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Contient un décalage, en caractères, à utiliser pour le rendu d’une pièce jointe dans le texte du message principal.
  
|||
|:-----|:-----|
|Propriétés associées :  <br/> |PR_RENDERING_POSITION  <br/> |
|Identificateur :  <br/> |0x370B  <br/> |
|Type de données :  <br/> |PT_LONG  <br/> |
|Domaine :  <br/> |Pièce jointe MAPI  <br/> |
   
## <a name="remarks"></a>Remarques

Lorsque le décalage fourni est -1 (0xFFFFFFFF), la pièce jointe n’est pas rendue à l’aide de cette propriété. Toutes les valeurs autres que -1 indiquent la position dans la propriété **PR_BODY** ([PidTagBody](pidtagbody-canonical-property.md)) à laquelle la pièce jointe doit être rendue.
  
 **Remarque** Le caractère indiqué par cette propriété dans **PR_BODY** est remplacé par la pièce jointe. En règle générale, ce caractère est un espace, bien qu’un caractère d’espace réservé spécial puisse également être utilisé. 
  
Cette propriété est exprimée en caractères. Dans certains jeux de caractères, cela n’équivaut pas à des octets. Les applications Unicode peuvent calculer la position en fonction de caractères à deux caractères. Double-Byte applications de jeu de caractères (DBCS) doivent analyser le texte jusqu’à cette valeur de propriété, car leur représentation des caractères varie d’un à deux octets par caractère.
  
Cette propriété ne doit pas être utilisée avec du texte RTF (Rich Text Format). La position de rendu est indiquée au rtf par une séquence d’échappatoire appelée espace réservé de pièce jointe d’objet. Cette séquence se compose de la chaîne suivie d’un seul caractère, normalement un espace, qui sera remplacé par le rendu  `\objattph` des pièces jointes. 
  
## <a name="related-resources"></a>Ressources connexes

### <a name="protocol-specifications"></a>Spécifications de protocole

[[MS-OXPROPS]](https://msdn.microsoft.com/library/f6ab1613-aefe-447d-a49c-18217230b148%28Office.15%29.aspx)
  
> Fournit des références aux spécifications Exchange Server protocole.
    
[[MS-OXCMSG]](https://msdn.microsoft.com/library/7fd7ec40-deec-4c06-9493-1bc06b349682%28Office.15%29.aspx)
  
> Gère les objets message et pièce jointe.
    
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

