---
title: Fichiers et messages joints
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
localization_priority: Normal
api_type:
- COM
ms.assetid: b2f2fb72-23ae-4e0b-a8a1-3b78a1862acb
description: 'Derniére modification : samedi 23 juillet 2011'
ms.openlocfilehash: c4dbe1a761a753bef77168aec8d2674a1b2b100e
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33436838"
---
# <a name="attached-files-and-messages"></a>Fichiers et messages joints

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Si MIME avec TNEF est utilisé lors du codage du contenu des messages, toutes les propriétés et le contenu des pièces jointes sont dans le flux TNEF. Le format TNEF lui-même est un fichier binaire joint unique nommé Winmail.dat, codé comme décrit pour MIME sans TNEF. 
  
Si MIME est utilisé sans TNEF, les fichiers joints sont envoyés en tant que composants de contenu de message MIME. Le nom de fichier est placé dans le paramètre  *name*  de  *l’en-tête Content-type*  de la pièce jointe. Le jeu de caractères de la pièce jointe est placé dans le  *paramètre charset*  au  *type de contenu*  ; Il et le contenu-transfer-encoding sont déterminés par l’analyse de l’intégralité du contenu de la pièce jointe. Les pièces jointes d’URL sont traitées de manière spéciale : 
  
- Si la pièce jointe est une URL (un fichier joint avec une extension . URL) et le mode d’accès défini dans ce fichier est FTP anonyme, il est codé en tant que message externe et le contenu du fichier (l’URL) est copié dans l’en-tête du message externe. *Content-type: message/external-body; access-type=anon-ftp*  (Content-Transfer-Encoding: 7bit is assumed.) 
    
- Si seuls des caractères 7 bits sont trouvés et qu’aucune ligne ne dépasse 140 caractères, la pièce jointe est du texte ASCII. *Content-type: text/plain; charset=us-ascii Content-Transfer-Encoding: 7bit* 
    
- Si des lignes longues ou jusqu’à 25 % de caractères 8 bits sont trouvés, le contenu de la pièce jointe est du texte et le jeu de caractères est déterminé par les paramètres régionaux. Il doit être choisi parmi les jeux de caractères définis par la norme ISO 8859. *Content-type: text/plain; charset=ISO-8859-1*  (par exemple) 
    
     *Content-Transfer-Encoding: quoted-printable* 
    
- Si 25 % ou plus des caractères ont le jeu de bits élevé, la pièce jointe est binaire. Il est codé à l’aide de l’algorithme Base64. *Content-type: application/octet-stream*  (par défaut ; en fonction de l’extension de fichier) 
    
     * Content-Transfer-Encoding: base64 * 
    
Sur les messages sortants, le type de contenu doit être dérivé de l’extension à trois lettres du nom de fichier. Ce mappage existe dans le Registre système ; sous se trouve une valeur de chaîne nommée « Content Type » qui donne au type de contenu MIME s’il est défini. Cet exemple est pour un fichier image TIFF :
  
HKEY_LOCAL_MACHINE\
  
Software\
  
Microsoft\
  
Classes\
  
.tif
  
Type de contenu = « image/tiff »
  
S’il n’existe aucun mappage pour l’extension de fichier, le flux  *d’application/d’octet*  par défaut doit être utilisé. 
  
Sur les messages entrants, le type de contenu d’une pièce jointe doit toujours être copié dans la propriété MAPI **PR_ATTACH_MIME_TAG** ([PidTagAttachMimeTag](pidtagattachmimetag-canonical-property.md)). Même si un nom de fichier est défini pour un fichier joint, l’extension mappée par le type de contenu doit être utilisée dans les propriétés **PR_ATTACH_FILENAME** ([PidTagAttachFilename](pidtagattachfilename-canonical-property.md)) et **PR_ATTACH_EXTENSION** ([PidTagAttachExtension](pidtagattachextension-canonical-property.md)).
  
Le  *paramètre*  name est officiellement supprimé par la RFC 821. Au fil de l’évolution des normes, Microsoft envisagera de spécifier un mappage de remplacement pour les noms de fichiers joints. 
  
Les messages joints sortants sont envoyés sous la forme * Type de contenu : message/rfc822 * Les messages joints sont codés de manière récursive, à leur place. Les composants de contenu de message entrant avec  *Content-Type: multipart/digest*  sont également mappés sur des messages incorporés. 
  
Si uuencode avec TNEF est utilisé lors du codage du contenu des messages, toutes les propriétés et le contenu des pièces jointes se trouve dans le flux TNEF. Le format TNEF lui-même est un fichier binaire joint unique nommé Winmail.dat, codé comme décrit pour Uuencode sans TNEF.
  
Si uuencode est utilisé sans TNEF, tous les fichiers joints sont traités comme binaires et uuencoded, à la suite du texte du message. Le nom de fichier est présent dans l’en-tête uuencode :
  
 begin 0755 Winmail.dat 
  
 ... données ... 
  
 end 
  
Les messages joints sont textuels dans le texte du message. La hiérarchie des messages joints est toujours aplatie ; autrement dit, les messages au sein des messages joints sont retirés au niveau supérieur.
  
Les objets OLE incorporés sont ignorés.
  
Les positions de rendu des pièces jointes sont transmises littéralement, à l’aide de la **propriété PR_ATTACH_RENDERING** ([PidTagAttachRendering](pidtagattachrendering-canonical-property.md)) dans le TNEF. Si le TNEF n’est pas utilisé, ils sont perdus. Les pièces jointes entrantes sans position de rendu (y compris en l’absence de TNEF) ont leur position de rendu définie sur 0xFFFFFFFF, c’est-à-dire, sans position dans le texte du message.
  
## <a name="see-also"></a>Voir aussi



[Mappage des attributs de messagerie Internet aux propriétés MAPI](mapping-of-internet-mail-attributes-to-mapi-properties.md)

