---
title: IMsgServiceAdminAdminProviders
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IMsgServiceAdmin.AdminProviders
api_type:
- COM
ms.assetid: 0d605e2c-10db-46e1-95d5-12fabd524baa
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 6b7360995a781824b50ff02b5d2dec8e481e7ba7
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33422760"
---
# <a name="imsgserviceadminadminproviders"></a>IMsgServiceAdmin::AdminProviders

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Renvoie un pointeur qui donne accès à un objet d’administration de fournisseur.
  
```cpp
HRESULT AdminProviders(
  LPMAPIUID lpUID,
  ULONG ulFlags,
  LPPROVIDERADMIN FAR * lppProviderAdmin
);
```

## <a name="parameters"></a>Parameters

 _lpUID_
  
> [in] Pointeur vers la structure [MAPIUID](mapiuid.md) qui contient l’identificateur unique du service de message à administrer. 
    
 _ulFlags_
  
> [in] Toujours NULL. 
    
 _lppProviderAdmin_
  
> [out] Pointeur vers un pointeur vers un objet d’administration de fournisseur.
    
## <a name="return-value"></a>Valeur renvoyée

S_OK 
  
> L’objet d’administration du fournisseur a été renvoyé avec succès.
    
MAPI_E_NOT_FOUND 
  
> Le **MAPIUID pointé** par  _lpUID_ n’existe pas. 
    
## <a name="remarks"></a>Remarques

La **méthode IMsgServiceAdmin::AdminProviders** permet d’accéder à un objet d’administration de fournisseur. L’administration d’un fournisseur est un objet qui prend en charge l’interface [IProviderAdmin](iprovideradminiunknown.md) et permet aux clients de : 
  
- Ajouter des fournisseurs de services à un service de messagerie.
    
- Supprimer des fournisseurs de services d’un service de messagerie.
    
- Ouvrir les sections de profil.
    
- Accéder à la table des fournisseurs de services de messagerie.
    
Les types de modifications qui peuvent réellement être apportées à un service de message pendant que le profil est en cours d’utilisation dépendent du service de message. Toutefois, la plupart des services de messagerie ne supportent pas les modifications telles que l’ajout et la suppression de fournisseurs pendant l’utilisation du profil.
  
## <a name="notes-to-callers"></a>Remarques pour les appelants

Pour récupérer la structure **MAPIUID** pour le service de message à administrer, récupérez la colonne de propriété **PR_SERVICE_UID** ([PidTagServiceUid](pidtagserviceuid-canonical-property.md)) à partir de la ligne du service de message dans la table de service de message. Pour plus d’informations, voir la procédure décrite dans la méthode [IMsgServiceAdmin::CreateMsgService.](imsgserviceadmin-createmsgservice.md) 
  
## <a name="mfcmapi-reference"></a>Référence MFCMAPI

Pour voir un exemple de code MFCMAPI, consultez le tableau suivant.
  
|**Fichier**|**Fonction**|**Commentaire**|
|:-----|:-----|:-----|
|MsgServiceTableDlg.cpp  <br/> |CMsgServiceTableDlg::OnDisplayItem  <br/> |MFCMAPI utilise la méthode **IMsgServiceAdmin::AdminProviders** pour ouvrir un objet d’administration de fournisseur pour un service.  <br/> |
   
## <a name="see-also"></a>Voir aussi



[IProviderAdmin : IUnknown](iprovideradminiunknown.md)
  
[MAPIUID](mapiuid.md)
  
[IMsgServiceAdmin : IUnknown](imsgserviceadminiunknown.md)


[MFCMAPI comme un exemple de Code](mfcmapi-as-a-code-sample.md)

