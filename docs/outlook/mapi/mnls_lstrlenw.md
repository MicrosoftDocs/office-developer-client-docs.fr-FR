---
title: MNLS_lstrlenW
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: d342a956-1164-4c9c-b0bb-7a0b72dc97fc
description: 'Last modified: February 21, 2012'
ms.openlocfilehash: 31f699d1193e55a88e57a0f491658e0d537ef75d
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32338464"
---
# <a name="mnls_lstrlenw"></a>MNLS_lstrlenW

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Détermine la longueur de la chaîne Unicode spécifiée, à l’exception du caractère null de fin.
  
> [!TIP]
> Envisagez plutôt [d’utiliser StringCchLength.](https://msdn.microsoft.com/library/ms647539%28VS.85%29.aspx) 
  
```cpp
int MNLS_lstrlen(
  LPCWSTR lpsz);
```

## <a name="parameters"></a>Parameters

 _lpsz_
  
> [in] Chaîne Unicode terminée par null à vérifier.
    
## <a name="return-value"></a>Valeur renvoyée

La fonction renvoie un nombre integer avec la longueur de la chaîne. Il s’agit d’un nombre de caractères dans la chaîne, à l’exclusion du caractère null de fin. Si  _lpsz est_ NULL, la fonction renvoie zéro. 
  
## <a name="remarks"></a>Remarques

Cette fonction encapsule **la fonction lstrlen.** Pour plus d’informations, voir [lstrlen](https://msdn.microsoft.com/library/ms647492%28VS.85%29.aspx).
  
## <a name="see-also"></a>Voir aussi



[lstrlen](https://msdn.microsoft.com/library/ms647492%28VS.85%29.aspx)
  
[StringCchLength](https://msdn.microsoft.com/library/ms647539%28VS.85%29.aspx)

