---
title: IMessageDeleteAttach
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IMessage.DeleteAttach
api_type:
- COM
ms.assetid: 0a5cb49f-c4f3-4893-8616-80d6332efcfc
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: c14ccac0addbc1288e3507cf549344f75e69cc28
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33430706"
---
# <a name="imessagedeleteattach"></a>IMessage::DeleteAttach

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Supprime une pièce jointe.
  
```cpp
HRESULT DeleteAttach(
ULONG ulAttachmentNum,
ULONG_PTR ulUIParam,
LPMAPIPROGRESS lpProgress,
ULONG ulFlags
);
```

## <a name="parameters"></a>Parameters

 _ulAttachmentNum_
  
> [in] Numéro d’index de la pièce jointe à supprimer. Il s’agit de la valeur de la propriété **PR_ATTACH_NUM** ([PidTagAttachNumber](pidtagattachnumber-canonical-property.md)).
    
 _ulUIParam_
  
> [in] Handle vers la fenêtre parente de toutes les boîtes de dialogue ou fenêtres affichées par cette méthode. Le _paramètre ulUIParam_ est ignoré, sauf si l’ATTACH_DIALOG est définie dans le _paramètre ulFlags._ 
    
 _lpProgress_
  
> [in] Pointeur vers un objet de progression qui affiche un indicateur de progression. Si NULL est transmis dans  _lpProgress,_ le fournisseur de magasin de messages affiche un indicateur de progression à l’aide de l’implémentation de l’objet de progression MAPI. Le  _paramètre lpProgress est_ ignoré, sauf si l’ATTACH_DIALOG est définie dans  _ulFlags_.
    
 _ulFlags_
  
> [in] Masque de bits d’indicateurs qui contrôle l’affichage d’une interface utilisateur. L’indicateur suivant peut être définie :
    
ATTACH_DIALOG 
  
> Demande l’affichage d’un indicateur de progression au cours de l’opération.
    
## <a name="return-value"></a>Valeur renvoyée

S_OK 
  
> La pièce jointe a été supprimée avec succès.
    
## <a name="remarks"></a>Remarques

La **méthode IMessage::D eleteAttach** supprime une pièce jointe d’un message. 
  
Une pièce jointe supprimée n’est pas définitivement supprimée tant que la méthode [IMAPIProp::SaveChanges](imapiprop-savechanges.md) du message n’a pas été appelée. 
  
## <a name="notes-to-callers"></a>Remarques pour les appelants

Avant **d’appeler DeleteAttach,** appelez la méthode **IUnknown::Release** pour la pièce jointe et chacun de ses flux. 
  
Étant donné que la suppression d’une pièce jointe peut être un processus long, **DeleteAttach** fournit le mécanisme qui affiche un indicateur de progression. Vous pouvez demander l’affichage d’un indicateur de progression en passant un pointeur vers votre [implémentation IMAPIProgress : IUnknown](imapiprogressiunknown.md) ou NULL si vous n’avez pas d’implémentation. Vous devez également spécifier une poignée de fenêtre dans le paramètre _ulUIParam_ et l’indicateur ATTACH_DIALOG dans le _paramètre ulFlags._ 
  
## <a name="mfcmapi-reference"></a>Référence MFCMAPI

Pour voir un exemple de code MFCMAPI, consultez le tableau suivant.
  
|**Fichier**|**Fonction**|**Commentaire**|
|:-----|:-----|:-----|
|AttachmentsDlg.cpp  <br/> |CAttachmentsDlg::OnDeleteSelectedItem  <br/> |MFCMAPI utilise **la méthode IMessage::D eleteAttach** pour supprimer la pièce jointe sélectionnée.  <br/> |
   
## <a name="see-also"></a>Voir aussi



[IMAPIProp::SaveChanges](imapiprop-savechanges.md)
  
[IMessage : IMAPIProp](imessageimapiprop.md)


[MFCMAPI comme un exemple de Code](mfcmapi-as-a-code-sample.md)

