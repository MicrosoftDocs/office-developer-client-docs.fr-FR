---
title: HrGetOneProp
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- HrGetOneProp
api_type:
- HeaderDef
ms.assetid: 8d0a381a-e714-4663-9a57-b0e1cdbd6ba7
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 95273bf6025d6ef995d7c21c0e44bdbbf59072f6
ms.sourcegitcommit: fb521c23df785c9c3aefa5062272b2630a32e587
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/20/2021
ms.locfileid: "52589172"
---
# <a name="hrgetoneprop"></a>HrGetOneProp

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Extrait la valeur d’une propriété unique à partir d’une interface de propriété, c’est-à-dire une interface dérivée [d’IMAPIProp](imapipropiunknown.md). 
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapiutil.h  <br/> |
|Implémenté par :  <br/> |MAPI  <br/> |
|Appelé par :  <br/> |Applications clientes et fournisseurs de services  <br/> |
   
```cpp
HRESULT HrGetOneProp(
  LPMAPIPROP pmp,
  ULONG ulPropTag,
  LPSPropValue FAR * ppprop
);
```

## <a name="parameters"></a>Parameters

 _pmp_
  
> [in] Pointeur vers [l’interface IMAPIProp](imapipropiunknown.md) à partir de laquelle la valeur de propriété doit être récupérée. 
    
 _ulPropTag_
  
> [in] Balise de propriété de la propriété à récupérer. 
    
 _ppprop_
  
> [out] Pointeur vers un pointeur vers la structure [SPropValue renvoyée](spropvalue.md) définissant la valeur de propriété récupérée. 
    
## <a name="return-value"></a>Valeur renvoyée

MAPI_E_NOT_FOUND 
  
> La propriété demandée n’est pas disponible à partir de l’interface spécifiée.
    
## <a name="remarks"></a>Remarques

Contrairement à [la méthode IMAPIProp::GetProps,](imapiprop-getprops.md) la fonction **HrGetOneProp** ne renvoie jamais d’avertissement. Étant donné qu’elle récupère une seule propriété, elle réussit ou échoue simplement. Pour récupérer plusieurs propriétés, **GetProps** est plus rapide. 
  
Vous pouvez définir ou modifier une propriété unique avec [la fonction HrSetOneProp.](hrsetoneprop.md) 
  
## <a name="mfcmapi-reference"></a>Référence MFCMAPI

Pour voir un exemple de code MFCMAPI, consultez le tableau suivant.
  
|**Fichier**|**Fonction**|**Commentaire**|
|:-----|:-----|:-----|
|MAPIFunctions.cpp  <br/> |GetMAPIObjectType  <br/> |MFCMAPI utilise la **méthode HrGetOneProp** pour récupérer le type d’un objet.  <br/> |
   
## <a name="see-also"></a>Voir aussi



[MFCMAPI comme un exemple de Code](mfcmapi-as-a-code-sample.md)

