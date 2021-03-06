---
title: ISocialProviderVersion
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: dfc92878-ab8b-4721-aee8-997c56a8e45b
description: Renvoie une chaîne qui représente le numéro de version du fournisseur pour ce réseau social.
ms.openlocfilehash: 0749b8fb83a11328233442b79e1f9de50076effe
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33438273"
---
# <a name="isocialproviderversion"></a>ISocialProvider::Version

Renvoie une chaîne qui représente le numéro de version du fournisseur pour ce réseau social. 
  
```cpp
[propget] HRESULT _stdcall Version([out, retval] BSTR* Version);
```

## <a name="property-value"></a>Valeur de la propriété

Chaîne qui contient le numéro de version du fournisseur.
  
## <a name="remarks"></a>Remarques

La chaîne de version doit utiliser  _la version MajeureVersion_. _Format MinorVersion_ (par exemple, 1,4730). 
  
## <a name="see-also"></a>Voir aussi

- [ISocialProvider : IUnknown](isocialprovideriunknown.md)

