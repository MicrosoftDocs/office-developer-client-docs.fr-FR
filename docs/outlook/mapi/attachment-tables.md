---
title: Tables des pièces jointes
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
localization_priority: Normal
api_type:
- COM
ms.assetid: 92a07f7b-d34c-4085-ab11-eadcd918fa1b
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 4aa800b504e7ffb07d94ace6d8dc30c1463ed637
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33427442"
---
# <a name="attachment-tables"></a>Tables des pièces jointes

**S’applique à** : Outlook 2013 | Outlook 2016 
  
Une table de pièces jointes contient des informations sur tous les objets de pièce jointe associés à un message envoyé ou à un message en cours de composition. 
  
Seules les pièces jointes qui ont été enregistrées par le biais d’un appel à la méthode [IMAPIProp::SaveChanges](imapiprop-savechanges.md) du message sont incluses dans le tableau. Les tables de pièces jointes sont implémentées par les fournisseurs de magasins de messages et utilisées par les applications clientes et les fournisseurs de transport. 
  
Une table de pièces jointes est accessible en appelant l’une des opérations suivantes :
  
- [IMessage::GetAttachmentTable](imessage-getattachmenttable.md)
    
- [IMAPIProp::OpenProperty](imapiprop-openproperty.md), requesting the **PR_MESSAGE_ATTACHMENTS** ([PidTagMessageAttachments](pidtagmessageattachments-canonical-property.md)) property.
    
Les tables de pièces jointes sont dynamiques.
  
Les fournisseurs de magasins de messages ne sont pas tenus de prendre en charge le tri sur leurs tables de pièces jointes. Si le tri n’est pas pris en charge, le tableau doit être présenté dans l’ordre en fonction de la position de rendu ( **PR_RENDERING_POSITION** ([PidTagRenderingPosition](pidtagrenderingposition-canonical-property.md)).
  
Les fournisseurs de magasins de messages ne sont pas non plus tenus de prendre en charge les restrictions sur leurs tables de pièces jointes. Les fournisseurs qui ne viennent pas en charge les restrictions MAPI_E_NO_SUPPORT à partir de leurs implémentations [d’IMAPITable::Restrict](imapitable-restrict.md) et [IMAPITable::FindRow](imapitable-findrow.md).
  
Les tables de pièces jointes peuvent être petites ; il n’y a que quatre colonnes dans l’ensemble de colonnes requis :
  
- **PR_ATTACH_NUM** ([PidTagAttachNumber](pidtagattachnumber-canonical-property.md)) 
    
- **PR_INSTANCE_KEY** ([PidTagInstanceKey](pidtaginstancekey-canonical-property.md)) 
    
- **PR_RECORD_KEY** ([PidTagRecordKey](pidtagrecordkey-canonical-property.md)) 
    
- **PR_RENDERING_POSITION**
    
 **PR_ATTACH_NUM** est nontransmitable et contient une valeur permettant d’identifier de manière unique une pièce jointe dans un message. Cette propriété est souvent utilisée comme index dans les lignes du tableau. **PR_ATTACH_NUM** a une courte durée de vie ; il n’est valide que lorsque le message contenant la pièce jointe est ouvert. Sa valeur reste constante tant que la table des pièces jointes est ouverte. 
  
 **PR_INSTANCE_KEY** est obligatoire dans presque tous les tableaux. Il est utilisé pour identifier de manière unique une ligne particulière. 
  
 **PR_RECORD_KEY** est généralement utilisé pour identifier de manière unique un objet à des fins de comparaison. Contrairement **PR_ATTACH_NUM**, **PR_RECORD_KEY** a la même étendue qu’un identificateur d’entrée à long terme ; il reste disponible et valide même après la fermeture et la réouverture du message. Pour plus d’informations sur l’utilisation des clés d’enregistrement dans MAPI, voir [MapI Record and Search Keys](mapi-record-and-search-keys.md).
  
 **PR_RENDERING_POSITION** indique comment une pièce jointe doit être affichée dans un message texte enrichi. Il peut être décalé en caractères, le premier caractère du contenu du message étant stocké dans la propriété **PR_BODY** ([PidTagBody](pidtagbody-canonical-property.md)) étant décalé de 0 ou de -1 (0xFFFFFFFF), ce qui indique que la pièce jointe ne doit pas être rendue du tout dans le texte du message. Ne pas **définir PR_RENDERING_POSITION** est également une option. 
  
Lorsqu’une table de pièces jointes est triée par position de rendu, le fournisseur de la boutique de messages la traite comme une valeur signée (PT_LONG). Par conséquent, les pièces jointes dont les positions de rendu sont de -1 sont triées avant les pièces jointes dont les positions de rendu reflètent des décalages valides. 
  
Pour plus d’informations sur le rendu d’une pièce jointe dans un message en texte simple, voir [Rendu d’une pièce jointe en texte simple.](rendering-an-attachment-in-plain-text.md) 
  
Pour plus d’informations sur le rendu d’une pièce jointe dans du texte formaté tel que RTF (Rich Text Format), voir [Rendering an Attachment in RTF Text](rendering-an-attachment-in-rtf-text.md).
  
Certaines des propriétés que les fournisseurs de magasins de messages incluent généralement dans une table de pièces jointes, car elles sont faciles à calculer ou à récupérer sont les suivantes :
  
|||
|:-----|:-----|
|**PR_ATTACH_ENCODING** ([PidTagAttachEncoding](pidtagattachencoding-canonical-property.md))  <br/> |**PR_ATTACH_EXTENSION** ([PidTagAttachExtension](pidtagattachextension-canonical-property.md))  <br/> |
|**PR_ATTACH_FILENAME** ([PidTagAttachFilename](pidtagattachfilename-canonical-property.md))  <br/> |**PR_ATTACH_LONG_FILENAME** ([PidTagAttachLongFilename](pidtagattachlongfilename-canonical-property.md))  <br/> |
|**PR_ATTACH_PATHNAME** ([PidTagAttachPathname](pidtagattachpathname-canonical-property.md))  <br/> |**PR_ATTACH_LONG_PATHNAME** ([PidTagAttachLongPathname](pidtagattachlongpathname-canonical-property.md))  <br/> |
|**PR_ATTACH_METHOD** ([PidTagAttachMethod](pidtagattachmethod-canonical-property.md))  <br/> |**PR_ATTACH_TAG** ([PidTagAttachTag](pidtagattachtag-canonical-property.md))  <br/> |
|**PR_CREATION_TIME** ([PidTagCreationTime](pidtagcreationtime-canonical-property.md))  <br/> |**PR_ATTACH_TRANSPORT_NAME** ([PidTagAttachTransportName](pidtagattachtransportname-canonical-property.md))  <br/> |
|**PR_DISPLAY_NAME** ([PidTagDisplayName](pidtagdisplayname-canonical-property.md))  <br/> |**PR_LAST_MODIFICATION_TIME** ([PidTagLastModificationTime](pidtaglastmodificationtime-canonical-property.md))  <br/> |
   
## <a name="see-also"></a>Voir aussi

- [MAPI Tables](mapi-tables.md)

