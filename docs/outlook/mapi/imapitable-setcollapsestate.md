---
title: IMAPITableSetCollapseState
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IMAPITable.SetCollapseState
api_type:
- COM
ms.assetid: 31325e8f-1cf9-49b2-8118-953996b0037f
description: 'Derniére modification : samedi 23 juillet 2011'
ms.openlocfilehash: 7351457dc5b72cfc4a7ef9f91e9d33a80ca98c39
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33414066"
---
# <a name="imapitablesetcollapsestate"></a>IMAPITable::SetCollapseState

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Reconstruit l’état étendu ou réduire actuel d’une table classée à l’aide de données enregistrées par un appel précédent à la méthode [IMAPITable::GetCollapseState.](imapitable-getcollapsestate.md) 
  
```cpp
HRESULT SetCollapseState(
ULONG ulFlags,
ULONG cbCollapseState,
LPBYTE pbCollapseState,
BOOKMARK FAR * lpbkLocation
);
```

## <a name="parameters"></a>Parameters

 _ulFlags_
  
> Réservé ; doit être zéro.
    
 _cbCollapseState_
  
> [in] Nombre d’octets dans la structure pointée par le _paramètre pbCollapseState._ 
    
 _pbCollapseState_
  
> [in] Pointeur vers les structures contenant les données nécessaires à la reconstruction de la vue de table.
    
 _lpbkLocation_
  
> [out] Pointeur vers un signet identifiant la ligne du tableau à partir de laquelle l’état réduire ou développé doit être reconstruit. Ce signet et la clé d’instance transmises dans le paramètre  _lpbInstanceKey_ dans l’appel à [IMAPITable::GetCollapseState](imapitable-getcollapsestate.md) identifient la même ligne. 
    
## <a name="return-value"></a>Valeur renvoyée

S_OK 
  
> L’état de la table classée a été correctement reconstruit.
    
MAPI_E_BUSY 
  
> Une autre opération est en cours qui empêche le démarrage de l’opération. L’opération en cours doit être autorisée ou arrêtée.
    
MAPI_E_UNABLE_TO_COMPLETE 
  
> Le tableau n’a pas pu terminer la reconstruction de la vue collapsed ou expanded.
    
## <a name="remarks"></a>Remarques

La **méthode IMAPITable::SetCollapseState** rétablit l’état développé ou réduire de l’affichage tableau. **SetCollapseState et** **GetCollapseState** fonctionnent ensemble comme suit : 
  
1. Lorsque l’état d’une table classée est sur le point de changer, [IMAPITable::GetCollapseState](imapitable-getcollapsestate.md) est appelé pour enregistrer toutes les données relatives à l’état avant la modification. 
    
2. Pour restaurer l’état enregistré de la table, **SetCollapseState** est appelé. Les données enregistrées **par GetCollapseState** sont **transmises à SetCollapseState**. **SetCollapseState** est en mesure d’utiliser ces données pour restaurer l’état. 
    
3. **SetCollapseState** renvoie en tant que paramètre de sortie un signet qui identifie la même ligne que la clé d’instance transmise en tant qu’entrée à **GetCollapseState**.
    
Pour plus d’informations sur les tableaux classés, voir [Tri et catégorisation.](sorting-and-categorization.md) 
  
## <a name="notes-to-implementers"></a>Remarques pour les responsables de l’implémentation

Vous devez vérifier que l’ordre de tri et les restrictions sont exactement les mêmes qu’au moment de l’appel **getCollapseState.** Si une modification a été faite, **SetCollapseState** ne doit pas être appelé, car les résultats peuvent être imprévisibles. Cela peut se produire si, par exemple, un client appelle **GetCollapseState,** puis **SortTable** pour modifier la clé de tri avant d’appeler **SetCollapseState**. Pour des raisons de sécurité, vérifiez que les données enregistrées sont toujours valides avant de poursuivre la restauration. 
  
## <a name="notes-to-callers"></a>Remarques pour les appelants

Pour appeler **SetCollapseState,** vous devez avoir précédemment **appelé GetCollapseState**. L’ordre de tri qui établit les catégories doit être le même pour les deux méthodes. Si les ordres de tri diffèrent, les résultats de l’opération **SetCollapseState** sont imprévisibles. 
  
## <a name="see-also"></a>Voir aussi



[IMAPITable::CreateBookmark](imapitable-createbookmark.md)
  
[IMAPITable::FreeBookmark](imapitable-freebookmark.md)
  
[IMAPITable::GetCollapseState](imapitable-getcollapsestate.md)
  
[IMAPITable : IUnknown](imapitableiunknown.md)

