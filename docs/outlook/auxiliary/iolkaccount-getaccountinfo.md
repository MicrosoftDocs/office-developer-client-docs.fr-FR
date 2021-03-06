---
title: IOlkAccountGetAccountInfo
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: 97f08cde-d6e4-8935-1758-4018a3baf682
description: Obtient les informations de type et de catégorie pour le compte spécifié.
ms.openlocfilehash: 88021537cc7ff4c55759081e6f3619c2a9f10ea3
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33437902"
---
# <a name="iolkaccountgetaccountinfo"></a>IOlkAccount::GetAccountInfo

Obtient les informations de type et de catégorie pour le compte spécifié.
  
## <a name="quick-info"></a>Informations rapides

Voir [IOlkAccount](iolkaccount.md).
  
```cpp
HRESULT IOlkAccount::GetAccountInfo(  
    CLSID *pclsidType, 
    DWORD *pcCategories, 
    CLSID **prgclsidCategory 
);

```

## <a name="parameters"></a>Parameters

_pclsidType_
  
> [out] Identificateur de classe pour le type de compte. La valeur doit être une des opérations suivantes :
    
   - CLSID_OlkPOP3Account 
    
   - CLSID_OlkIMAP4Account 
    
   - CLSID_OlkMAPIAccount 
    
   - CLSID_OlkHotmailAccount 
    
   - CLSID_OlkLDAPAccount
    
_pcCategories_
  
> [out] Nombre de catégories dans  _prgclsidCategory_.
    
_prgclsidCategory_
  
> [out] Tableau des catégories associées à ce compte. Le tableau est de taille * _pcCategories_. La valeur de chaque catégorie du tableau doit être l’une des suivantes :
    
   - CLSID_OlkMail
    
   - CLSID_OlkAddressBook
    
   - CLSID_OlkStore
    
## <a name="return-values"></a>Valeurs de retour

S_OK si l'appel a réussi ; dans le cas contraire, un code d'erreur.
  
## <a name="remarks"></a>Remarques

Après le retour de cette méthode, vous devez libérer  *prgclsidCategory*  à l’aide de [IOlkAccount::FreeMemory](iolkaccount-freememory.md).
  
**IOlkAccount::GetAccountInfo** ne prend pas en charge la catégorie de carnet d’adresses pour Exchange compte. Si le compte est un compte Exchange (*pclsidType* est **CLSID_OlkMAPIAccount** ), et que le compte implémente le carnet d’adresses, l’appel de **IOlkAccount::GetAccountInfo** ne retournera pas **CLSID_OlkAddressBook** en tant que catégorie dans *prgclsidCategory* . 
  
## <a name="see-also"></a>Voir aussi

- [Constantes (API de gestion des comptes)](constants-account-management-api.md)  
- [IOlkAccount::FreeMemory](iolkaccount-freememory.md)

