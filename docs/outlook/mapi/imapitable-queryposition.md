---
title: IMAPITableQueryPosition
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IMAPITable.QueryPosition
api_type:
- COM
ms.assetid: 510b2e21-ba27-47dd-87cb-2a549e31fa28
description: 'Derniére modification : samedi 23 juillet 2011'
ms.openlocfilehash: 2e44d824bbb5cc96c51d7ca91eb639001bc52a71
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33415662"
---
# <a name="imapitablequeryposition"></a>IMAPITable::QueryPosition

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Extrait la position de ligne de tableau actuelle du curseur, en fonction d’une valeur fractionnaire.
  
```cpp
HRESULT QueryPosition(
ULONG FAR * lpulRow,
ULONG FAR * lpulNumerator,
ULONG FAR * lpulDenominator
);
```

## <a name="parameters"></a>Parameters

 _lpulRow_
  
> [out] Pointeur vers le numéro de la ligne en cours. Le numéro de ligne est de base zéro ; la première ligne du tableau est zéro. 
    
 _lpulNumerator_
  
> [out] Pointeur vers le numérateur de la fraction identifiant la position du tableau.
    
 _lpulDenominator_
  
> [out] Pointeur vers le dénominateur de la fraction identifiant la position du tableau. Le  _paramètre lpulDenominator ne peut_ pas être zéro. 
    
## <a name="return-value"></a>Valeur renvoyée

S_OK 
  
> La méthode a renvoyé des valeurs valides dans  _lpulRow,_  _lpulNumerator_ et  _lpulDenominator_.
    
## <a name="remarks"></a>Remarques

La **méthode IMAPITable::QueryPosition** détermine la position de ligne actuelle et renvoie le numéro de la ligne actuelle et une valeur fractionnaire indiquant sa position relative à la fin du tableau. MAPI définit la ligne actuelle comme la ligne suivante à lire. 
  
## <a name="notes-to-implementers"></a>Remarques pour les responsables de l’implémentation

Vous n’avez pas besoin de renvoyer le nombre exact de lignes dans la table pour le  _paramètre lpulDenominator_ ; Il peut s’agit d’une estimation. 
  
Si vous ne pouvez pas déterminer la ligne actuelle, renvoyez une valeur de 0xFFFFFFFF dans  _lpulRow_.
  
## <a name="notes-to-callers"></a>Remarques pour les appelants

Vous pouvez utiliser **QueryPosition** pour positionner une case de défilement dans une barre de défilement. Par exemple, dans un tableau contenant 100 lignes, si **QueryPosition** renvoie la valeur 75 dans le paramètre  _lpulNumerator,_ 100 dans le paramètre  _lpulDenominator_ et 75 dans le paramètre  _lpulRow,_ vous pouvez positionner la case de défilement 3/4 de la manière dans la barre de défilement. 
  
Ne comptez pas sur la valeur dans  _lpulDenominator_ qui est le nombre de lignes dans le tableau. **QueryPosition ne** peut pas toujours identifier la ligne exacte sur la position du curseur. 
  
Un appel à **QueryPosition** peut impliquer de grandes quantités de mémoire, en particulier pour les tables classées de grande taille. Si le  _paramètre lpulRow_ est 0xFFFFFFFF, trop de mémoire était nécessaire pour **que QueryPosition** détermine la ligne actuelle. Appelez [la méthode IMAPITable::SeekRowApprox](imapitable-seekrowapprox.md) pour positionner le tableau sur la ligne identifiée par les paramètres _lpulNumerator_ et _lpulDenominator._ Toutefois, n’attendez pas toujours que **SeekRowApprox** établisse en tant que position actuelle la même ligne **que QueryPosition** aurait si la mémoire n’avait pas été un facteur. 
  
## <a name="see-also"></a>Voir aussi



[IMAPITable::SeekRowApprox](imapitable-seekrowapprox.md)
  
[IMAPITable : IUnknown](imapitableiunknown.md)

