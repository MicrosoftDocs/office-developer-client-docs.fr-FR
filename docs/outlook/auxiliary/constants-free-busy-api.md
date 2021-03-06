---
title: Constantes (API de libre/occupé)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: overview
localization_priority: Normal
ms.assetid: ff756bf1-9395-5e50-4f55-1eb0365a84ed
description: Cette rubrique contient des définitions de constantes, des identificateurs de classe et des identificateurs d’interface pour l’API de libre/occupé.
ms.openlocfilehash: 13578617eaab45e7194d7a0d4d6995ef925e7f20
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33431532"
---
# <a name="constants-freebusy-api"></a>Constantes (API de libre/occupé)

Cette rubrique contient des définitions de constantes, des identificateurs de classe et des identificateurs d’interface pour l’API de libre/occupé.
  
## <a name="constants"></a>Constantes

|**Constante**|**Définition**|
|:-----|:-----|
|E_NOTIMPL  <br/> | *Comme défini dans le fichier d’en-tête du Kit de développement logiciel (SDK) microsoft Windows,winerror.h.*  <br/> |
|E_OUTOFMEMORY  <br/> | *Comme défini dans le fichier d Windows en-tête du SDK winerror.h.*  <br/> |
|S_FALSE  <br/> | *Comme défini dans le fichier d Windows en-tête du SDK winerror.h.*  <br/> |
|S_OK  <br/> | *Comme défini dans le fichier d Windows en-tête du SDK winerror.h.*  <br/> |
   
## <a name="interface-identifiers"></a>Identificateurs d’interface

Pour les identificateurs d’interface suivants, supposons que la macro DEFINE_GUID définie dans le fichier d’en-tête du SDK Windows guiddef.h associe le nom symbolique GUID à sa valeur.
  
{00067064-0000-0000-C000-000000000046}
  
DEFINE_GUID(IID_IEnumFBBlock, 0x00067064, 0x0, 0x0, 0xc0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46) ;
  
{00067066-0000-0000-C000-000000000046}
  
DEFINE_GUID(IID_IFreeBusyData, 0x00067066, 0x0, 0x0, 0xc0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46) ;
  
{00067067-0000-0000-C000-000000000046}
  
DEFINE_GUID(IID_IFreeBusySupport, 0x00067067, 0x0, 0x0, 0xc0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46) ;
  

