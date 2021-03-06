---
title: Utilisation de macros pour la gestion des erreurs
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
localization_priority: Normal
api_type:
- COM
ms.assetid: 351405ca-b72b-4e9e-bc8e-947344588970
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: 9e6901d6583e7a7924a4a7c19c0a262bcef74bd3
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33435564"
---
# <a name="using-macros-for-error-handling"></a>Utilisation de macros pour la gestion des erreurs

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Il existe plusieurs macros qui facilitent l’emploi des valeurs HRESULT.
  
Il existe deux ensembles de macros qui testent l’échec ou la réussite : HR_SUCCEEDED et HR_FAILED et SUCCEEDED et FAILED. SUCCEEDED est identique à HR_SUCCEEDED et FAILED est identique à HR_FAILED.
  
Dans ce cas, utilisez la macro **ResultFromScode** pour définir la variable HRESULT sur la valeur HRESULT correspondante pour S_OK. 
  
Les macros couramment utilisées sont brièvement décrites dans le tableau suivant.
  
|**Macro**|**Description**|
|:-----|:-----|
|**MAKE_HRESULT** <br/> |Construit un HRESULT à partir de ses composants.  <br/> |
|**HR_SUCCEEDED** <br/> |Teste un HRESULT pour une condition de réussite ou d’avertissement.  <br/> |
|**HR_FAILED** <br/> |Teste une condition d’erreur HRESULT.  <br/> |
|**HRESULT_CODE** <br/> |Extrait la partie code d’erreur du HRESULT.  <br/> |
|**HRESULT_FACILITY** <br/> |Extrait la fonction du HRESULT.  <br/> |
|**HRESULT_SEVERITY** <br/> |Extrait le bit de gravité de la GRAVITÉ.  <br/> |
|**SUCCEEDED** <br/> |Teste un HRESULT pour une condition de réussite ou d’avertissement.  <br/> |
|**FAILED** <br/> |Teste une condition d’erreur HRESULT.  <br/> |
   

