---
title: Plages d’identificateurs de propriétés
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
localization_priority: Normal
api_type:
- COM
ms.assetid: c01e95bb-be25-490d-880b-60674f890258
description: Dernière modification le 9 mars 2015
ms.openlocfilehash: cdf1eae945cddf9eb76a2b7a2532d5dc6568beac
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33422704"
---
# <a name="property-identifier-ranges"></a>Plages d’identificateurs de propriétés

  
  
**S’applique à** : Outlook 2013 | Outlook 2016 
  
Le tableau suivant récapitule les différentes plages d’identificateurs de propriété, décrivant le propriétaire des propriétés de chaque plage.
  
|**Plage d’identificateurs**|**Description**|
|:-----|:-----|
|0000  <br/> |Réservé par MAPI pour la valeur spéciale **PR_NULL**.  <br/> |
|0001 - 0BFF  <br/> |Propriétés d’enveloppe de message définies par MAPI.  <br/> |
|0C00 - 0DFF  <br/> |Propriétés de destinataire définies par MAPI.  <br/> |
|0E00 - 0FFF  <br/> |Propriétés de message non transmises définies par MAPI.  <br/> |
|1000 - 2FFF  <br/> |Propriétés de contenu de message définies par MAPI.  <br/> |
|3000 - 3FFF  <br/> |Propriétés des objets autres que les messages et les destinataires définis par MAPI.  <br/> |
|4000 - 57FF  <br/> |Propriétés d’enveloppe de message définies par les fournisseurs de transport.  <br/> |
|5800 - 5FFF  <br/> |Propriétés de destinataire définies par les fournisseurs de transport et de carnet d’adresses.  <br/> |
|6000 - 65FF  <br/> |Propriétés de message non transmises définies par les clients.  <br/> |
|6600 - 67FF  <br/> |Propriétés non transmises définies par un fournisseur de services. Ces propriétés peuvent être visibles ou invisibles pour les utilisateurs.  <br/> |
|6800 - 7BFF  <br/> |Propriétés de contenu de message pour les classes de message personnalisées définies par les créateurs de ces classes.  <br/> |
|7C00 - 7FFF  <br/> |Propriétés non transmises pour les classes de message personnalisées définies par les créateurs de ces classes.  <br/> |
|8000 - FFFE  <br/> |Propriétés définies par les clients et parfois les fournisseurs de services identifiés par leur nom via les méthodes [IMAPIProp::GetNamesFromIDs](imapiprop-getnamesfromids.md) et [IMAPIProp::GetIDsFromNames.](imapiprop-getidsfromnames.md)  <br/> |
|FFFF  <br/> |Réservé par MAPI pour la valeur d’erreur spéciale PROP_ID_INVALID.  <br/> |
   
La plage entre 3 000 et 3FFF est réservée aux propriétés qui ne sont pas liées aux messages ou aux destinataires. MAPI divise cette plage en sous-plages par types d’objet ; Le tableau suivant présente cette répartition supplémentaire. 
  
|**Plage d’identificateurs**|**Type de propriété**|
|:-----|:-----|
|3000 - 33FF  <br/> |Propriétés courantes qui apparaissent sur plusieurs objets, telles que **PR_DISPLAY_NAME** et **PR_ENTRYID**.  <br/> |
|3400 - 35FF  <br/> |Propriétés de la boutique de messages  <br/> |
|3600 - 36FF  <br/> |Propriétés de conteneur de dossiers et de carnets d’adresses  <br/> |
|3700 - 38FF  <br/> |Propriétés de pièce jointe  <br/> |
|3900 - 39FF  <br/> |Propriétés du carnet d’adresses  <br/> |
|3A00 - 3BFF  <br/> |Propriétés utilisateur de messagerie  <br/> |
|3C00 - 3CFF  <br/> |Propriétés de liste de distribution  <br/> |
|3D00 - 3DFF  <br/> |Propriétés de profil  <br/> |
|3E00 - 3FFF  <br/> |Propriétés de l’objet Status  <br/> |
   

