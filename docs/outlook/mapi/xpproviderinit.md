---
title: XPProviderInit
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.XPProviderInit
api_type:
- COM
ms.assetid: df6eacf4-1cf9-4c25-806f-f87c38dad597
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: ee0ff8d32436f71020be2cdc91d6677bd4ec8e43
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33428535"
---
# <a name="xpproviderinit"></a>XPProviderInit

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Initialise un fournisseur de transport pour l’opération.
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapispi.h  <br/> |
|Implémenté par :  <br/> |Fournisseurs de transport  <br/> |
|Appelé par :  <br/> |MAPI  <br/> |
   
```cpp
HRESULT XPProviderInit(
  HINSTANCE hInstance,
  LPMALLOC lpMalloc,
  LPALLOCATEBUFFER lpAllocateBuffer,
  LPALLOCATEMORE lpAllocateMore,
  LPFREEBUFFER lpFreeBuffer,
  ULONG ulFlags,
  ULONG ulMAPIVer,
  ULONG FAR * lpulProviderVer,
  LPXPPROVIDER FAR * lppXPProvider
);
```

## <a name="parameters"></a>Parameters

 _hInstance_
  
> [in] Instance de la bibliothèque de liens dynamiques (DLL) du fournisseur de transport que MAPI a utilisée lors du chargement de la DLL.
    
 _lpMalloc_
  
> [in] Pointeur vers un objet d’allocation de mémoire exposant l’interface OLE **IMalloc.** Le fournisseur de transport peut avoir besoin d’utiliser cette méthode d’allocation lors de l’utilisation de certaines interfaces telles que **IStream**. 
    
 _lpAllocateBuffer_
  
> [in] Pointeur vers la [fonction MAPIAllocateBuffer,](mapiallocatebuffer.md) à utiliser pour allouer de la mémoire. 
    
 _lpAllocateMore_
  
> [in] Pointeur vers la [fonction MAPIAllocateMore,](mapiallocatemore.md) à utiliser pour allouer de la mémoire supplémentaire. 
    
 _lpFreeBuffer_
  
> [in] Pointeur vers la [fonction MAPIFreeBuffer,](mapifreebuffer.md) à utiliser pour libérer de la mémoire. 
    
 _ulFlags_
  
> [in] Masque de bits d’indicateurs. L’indicateur suivant peut être définie :
    
MAPI_NT_SERVICE 
  
> Le fournisseur est chargé dans le contexte d’un service Windows, un type spécial de processus sans accès à une interface utilisateur. 
    
 _ulMAPIVer_
  
> [in] Numéro de version de l’interface du fournisseur de services (SPI) que Mapi.dll utilise. Pour le numéro de version actuel, voir le fichier d’en-tête Mapispi.h. 
    
 _lpulProviderVer_
  
> [out] Pointeur vers le numéro de version du SPI utilisé par ce fournisseur de transport. 
    
 _lppXPProvider_
  
> [out] Pointeur vers un pointeur vers l’objet fournisseur de transport initialisé.
    
## <a name="return-value"></a>Valeur renvoyée

S_OK 
  
> L'appel a r�ussi et a renvoy� la valeur attendue ou les valeurs. 
    
MAPI_E_VERSION 
  
> La version SPI utilisée par MAPI n’est pas compatible avec la spi utilisée par ce fournisseur.
    
## <a name="remarks"></a>Remarques

MAPI appelle la fonction de point d’entrée **XPProviderInit** pour initialiser un fournisseur de transport à la suite d’une logon client. **XPProviderInit est** appelé une fois pour chaque fournisseur de transport spécifié dans le profil du client. 
  
## <a name="notes-to-implementers"></a>Remarques pour les responsables de l’implémentation

Un fournisseur de transport doit implémenter **XPProviderInit** en tant que fonction de point d’entrée dans la DLL du fournisseur. L’implémentation doit être basée sur le prototype de fonction **XPPROVIDERINIT,** également spécifié dans Mapispi.h. MAPI définit **XPPROVIDERINIT** pour utiliser le type d’appel d’initialisation MAPI standard, STDMAPIINITCALLTYPE, qui entraîne **XPProviderInit** à respecter la convention d’appel CDECL. Un avantage de CDECL est que les appels peuvent être tentés même si le nombre de paramètres d’appel ne correspond pas au nombre de paramètres définis. 
  
Un fournisseur peut être initialisé plusieurs fois suite à l’apparition de plusieurs profils dans une utilisation simultanée ou de l’apparition de plusieurs fois dans le même profil. Étant donné que l’objet fournisseur contient du contexte, **XPProviderInit** doit retourner un autre objet fournisseur dans  _lppXPProvider_ pour chaque initialisation, même pour plusieurs initialisations dans le même processus. 
  
Le fournisseur de transport doit utiliser les fonctions pointées par  _lpAllocateBuffer,_  _lpAllocateMore_ et  _lpFreeBuffer_ pour la plupart des allocations de mémoire et de la déallocation. En particulier, le fournisseur doit utiliser ces fonctions pour allouer de la mémoire pour une utilisation par les applications clientes lors de l’appel d’interfaces d’objet telles que [IMAPIProp::GetProps](imapiprop-getprops.md) et [IMAPITable::QueryRows](imapitable-queryrows.md). Si le fournisseur s’attend également à utiliser l’allocation de mémoire OLE, il doit appeler la méthode **IUnknown::AddRef** de l’objet d’allocation pointé par le paramètre _lpMalloc._ 
  
Pour plus d’informations **sur l’écriture de XPProviderInit,** voir [Initialisation du fournisseur de transport.](initializing-the-transport-provider.md) Pour plus d’informations sur les fonctions de point d’entrée, voir [Implementing a Service Provider Entry Point Function](implementing-a-service-provider-entry-point-function.md). 
  
## <a name="see-also"></a>Voir aussi



[ABProviderInit](abproviderinit.md)
  
[IXPProvider : IUnknown](ixpprovideriunknown.md)
  
[MSProviderInit](msproviderinit.md)

