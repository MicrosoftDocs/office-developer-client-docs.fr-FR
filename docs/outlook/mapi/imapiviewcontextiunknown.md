---
title: IMAPIViewContext IUnknown
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IMAPIViewContext
api_type:
- COM
ms.assetid: d566ff39-92c1-4a14-85e5-1c406825f805
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: db0c375218755c3a28475e2ebce2d097fb789f75
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33406030"
---
# <a name="imapiviewcontext--iunknown"></a>IMAPIViewContext : IUnknown

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Gère un formulaire dans la visionneuse de formulaires d’une application cliente. 
  
|||
|:-----|:-----|
|Fichier d’en-tête :  <br/> |Mapiform.h  <br/> |
|Exposé par :  <br/> |Afficher les objets de contexte  <br/> |
|Implémenté par :  <br/> |Visionneuses de formulaires  <br/> |
|Appelé par :  <br/> |Objets de formulaires  <br/> |
|Identificateur d’interface :  <br/> |IID_IMAPIViewContext  <br/> |
|Type de pointeur :  <br/> |LPMAPIVIEWCONTEXT  <br/> |
   
## <a name="vtable-order"></a>Ordre des vtables

|||
|:-----|:-----|
|[SetAdviseSink](imapiviewcontext-setadvisesink.md) <br/> |Gère l’inscription d’un formulaire pour recevoir des notifications sur les modifications apportées à la visionneuse.  <br/> |
|[ActivateNext](imapiviewcontext-activatenext.md) <br/> |Active le message suivant ou précédent dans la visionneuse de formulaires.  <br/> |
|[GetPrintSetup](imapiviewcontext-getprintsetup.md) <br/> |Récupère les informations d’impression actuelles.  <br/> |
|[GetSaveStream](imapiviewcontext-getsavestream.md) <br/> |Extrait un flux à utiliser pour enregistrer le message actuel.  <br/> |
|[GetViewStatus](imapiviewcontext-getviewstatus.md) <br/> |Récupère l’état actuel de la visionneuse.  <br/> |
|[GetLastError](imapiviewcontext-getlasterror.md) <br/> |Renvoie une structure [MAPIERROR](mapierror.md) qui contient des informations sur l’erreur précédente qui s’est produite dans l’objet de contexte d’affichage.  <br/> |
   
## <a name="see-also"></a>Voir aussi



[Interfaces MAPI](mapi-interfaces.md)

