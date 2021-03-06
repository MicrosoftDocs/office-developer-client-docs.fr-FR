---
title: SizedDtblPage
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MAPI.SizedDtblPage
api_type:
- COM
ms.assetid: 47b2a69d-e902-429f-8b31-166b51aeaf7f
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: f14b8d7a9a73997f797f9cfa26a2e574222e839e
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33407444"
---
# <a name="sizeddtblpage"></a>SizedDtblPage

**S’applique à** : Outlook 2013 | Outlook 2016 
  
Crée une structure nommée qui inclut une structure [DTBLPAGE](dtblpage.md) pour décrire un contrôle de page à onglets, une étiquette d’une longueur spécifiée et une entrée de fichier d’aide d’une longueur spécifiée. 
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapidefs.h  <br/> |
|Structure connexe :  <br/> |**DTBLPAGE** <br/> |
   
```cpp
SizedDtblPage (n, n1, u)
```

## <a name="parameters"></a>Parameters

_n_
  
> Longueur de l’étiquette de l’onglet de page.
    
_n1_
  
> Longueur de l’entrée apparaissant dans le fichier Mapisvc.inf identifiant le fichier d’aide qui sera utilisé avec le contrôle de page à onglets.
    
_u_
  
> Nom de la nouvelle structure.
    
## <a name="remarks"></a>Remarques

La macro **SizedDtblPage** vous permet de définir un contrôle de page à onglets lorsque le nombre de caractères dans l’étiquette associée et l’entrée du fichier d’aide est connu. La nouvelle structure est créée avec les membres suivants : 
  
```cpp
DTBLPAGE dtblpage;
TCHAR lpszLabel[n];
TCHAR lpszComponent[n1];
```

Pour utiliser un pointeur vers la structure résultante de la macro **SizedDtblPage** en tant que pointeur de structure **DTBLPAGE,** effectuez la distribution suivante : 
  
```cpp
lpDtblPage = (LPDTBLPAGE) &SizedDtblPage;
```

## <a name="see-also"></a>Voir aussi

- [DTBLPAGE](dtblpage.md)
- [Macros liées aux structures](macros-related-to-structures.md)

