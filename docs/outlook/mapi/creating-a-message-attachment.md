---
title: Création d’une pièce jointe de message
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
localization_priority: Normal
api_type:
- COM
ms.assetid: 711b6765-7763-41ae-9ff8-61ca6ddd459d
description: 'Derniére modification : samedi 23 juillet 2011'
ms.openlocfilehash: 2bdba3574c962c825f45cd098efa1cba6e21a4ea
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33405995"
---
# <a name="creating-a-message-attachment"></a>Création d’une pièce jointe de message
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Une pièce jointe de message est des données supplémentaires, telles qu’un fichier, un autre message ou un objet OLE, que vous pouvez envoyer ou enregistrer avec un message. Chaque pièce jointe possède une collection de propriétés qui l’identifie et décrit son type et la façon dont elle est rendue. Comme les destinataires, les pièces jointes de message sont accessibles uniquement par le biais du message auquel elles appartiennent. Par conséquent, pour qu’une pièce jointe soit utilisable, son message doit être ouvert.
  
## <a name="create-a-message-attachment"></a>Créer une pièce jointe de message
  
1. Appelez la méthode [IMessage::CreateAttach](imessage-createattach.md) du message et passez NULL comme identificateur d’interface. **CreateAttach** renvoie un nombre qui identifie de manière unique la nouvelle pièce jointe dans le message. Le numéro de pièce jointe est stocké dans la propriété **PR_ATTACH_NUM** ([PidTagAttachNumber](pidtagattachnumber-canonical-property.md)) et n’est valide que si le message contenant la pièce jointe est ouvert.
    
2. Appelez [IMAPIProp::SetProps](imapiprop-setprops.md) pour définir **PR_ATTACH_METHOD** ([PidTagAttachMethod](pidtagattachmethod-canonical-property.md)) pour indiquer comment accéder à la pièce jointe. **PR_ATTACH_METHOD** est obligatoire. Définissez-la sur : 
    
   - ATTACH_BY_VALUE si la pièce jointe est binaire.
    
   - ATTACH_BY_REFERENCE, ATTACH_BY_REF_RESOLVE ou ATTACH_BY_REF_ONLY si la pièce jointe est un fichier.
    
   - ATTACH_EMBEDDED_MSG si la pièce jointe est un message.
    
   - ATTACH_OLE si la pièce jointe est un objet OLE.
    
3. Définissez la propriété de données de pièce jointe appropriée :
    
   - **PR_ATTACH_DATA_BIN** ([PidTagAttachDataBinary](pidtagattachdatabinary-canonical-property.md)) pour les données binaires et les objets OLE 1.
    
   - **PR_ATTACH_PATHNAME** ([PidTagAttachPathname](pidtagattachpathname-canonical-property.md)) pour les fichiers.
    
   - **PR_ATTACH_DATA_OBJ** ([PidTagAttachDataObject](pidtagattachdataobject-canonical-property.md)) pour les messages et les objets OLE 2.
    
4. Définissez **PR_ATTACH_RENDERING** ([PidTagAttachRendering](pidtagattachrendering-canonical-property.md)) pour contenir la représentation graphique de la pièce jointe pour le fichier ou les pièces jointes binaires. Ne la définissez pas pour les objets OLE, qui stockent les informations de rendu en interne, ou pour les messages joints. 
    
5. Définissez **PR_RENDERING_POSITION** ([PidTagRenderingPosition](pidtagrenderingposition-canonical-property.md)) pour indiquer où la pièce jointe doit être affichée. **PR_RENDERING_POSITION** s’applique uniquement aux clients qui définissent **PR_BODY** propriété. Si vous ne le **PR_RTF_COMPRESSED,** placez les informations d’espace réservé suivantes dans le flux compressé :
    
   `\objattph`

   Pour définir **PR_RENDERING_POSITION**, affectez soit un nombre qui représente un décalage ordinal en caractères, avec le premier caractère de **PR_BODY** étant 0, si vous avez besoin de savoir où dans le message la pièce jointe est restituer, ou 0xFFFFFFFF, si vous ne restituer les pièces jointes dans **PR_BODY**.
    
6. Définissez **PR_ATTACH_FILENAME** ([PidTagAttachFilename](pidtagattachfilename-canonical-property.md)) pour indiquer le nom court du fichier pour une pièce jointe et pr **\_ ATTACH_LONG_FILENAME** ([PidTagAttachLongFilename](pidtagattachlongfilename-canonical-property.md)) pour indiquer le nom du fichier tel qu’il est pris en charge sur une plateforme qui gère le format de nom de fichier long. Les deux propriétés sont facultatives. Toutefois, si vous **définissez PR_ATTACH_LONG_FILENAME**, définissez également **PR_ATTACH_FILENAME**. 
    
7. Définissez **PR_DISPLAY_NAME** ([PidTagDisplayName](pidtagdisplayname-canonical-property.md)) pour indiquer le nom de la pièce jointe qui peut apparaître dans une boîte de dialogue. PR_DISPLAY_NAME est facultatif. 
    
## <a name="set-pr_attach_data_bin"></a>Définir PR_ATTACH_DATA_BIN
  
1. Appelez [IMAPIProp::OpenProperty](imapiprop-openproperty.md) pour ouvrir la propriété avec l’interface **IStream.** 
    
2. Si un fichier contient les données et qu’il est ouvert ou si vous avez besoin d’un contrôle explicite sur la taille de la mémoire tampon, appelez **IStream::Write** dans une boucle pour placer les données dans le flux. 
    
3. Une autre option consiste à appeler **OpenStreamOnFile** pour créer un flux pour accéder au fichier de données, puis appeler la méthode **IStream::CopyTo** de ce flux pour copier les données dans le flux renvoyé par **OpenProperty**.
    
4. Appelez la méthode **IStream::Commit** du nouveau flux. 
    
## <a name="set-pr_attach_data_obj"></a>Définir PR_ATTACH_DATA_OBJ
  
1. Appelez **IMAPIProp::OpenProperty** pour ouvrir la propriété avec l’interface **IStreamDocfile** afin de créer un flux qui fonctionne avec un stockage structuré. **IStreamDocfile** est implémenté par les fournisseurs de magasins de messages pour offrir aux clients un moyen plus performant de stocker et récupérer un stockage structuré. **L’interface IStreamDocfile** est identique à **IStream,** mais le contenu du flux est assuré d’être formaté en tant que stockage structuré. Si cet appel réussit, créez le flux avec les mêmes étapes décrites pour définir **PR_ATTACH_DATA_BIN**.
    
2. En cas **d’échec d’OpenProperty** : 
    
   1. Appelez **à nouveau OpenProperty** en demandant **IStorage**. 
      
   2. Appelez **StgOpenStorage pour ouvrir** l’objet OLE et renvoyer un objet de stockage. 
      
   3. Appelez la méthode **IStorage::CopyTo** de l’objet de stockage renvoyé pour copier dans l’objet de stockage renvoyé par **OpenProperty**.
      
   4. Appelez la méthode **IStorage::Commit** du nouvel objet de stockage. 
    
## <a name="set-pr_attach_pathname"></a>Définir PR_ATTACH_PATHNAME
  
1. Allouez une structure [SPropValue,](spropvalue.md) en fixant le membre **ulPropTag** sur **PR_ATTACH_PATHNAME** et le membre **Value.LPSZ** sur la chaîne de caractères qui représente le nom du fichier. 
    
2. Appelez la méthode [IMAPIProp::SetProps](imapiprop-setprops.md) de la pièce jointe. 
    
> [!NOTE]
> Si votre plateforme prend en charge les noms de fichiers longs, PR_ATTACH_PATHNAME **et** **PR_ATTACH_LONG_PATHNAME**. Il peut être nécessaire d’effectuer un appel de système d’exploitation pour récupérer le nom de fichier court. 
  
## <a name="see-also"></a>Voir aussi

- [Pièces jointes MAPI](mapi-attachments.md)

