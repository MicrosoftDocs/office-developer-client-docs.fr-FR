---
title: XML pour les fonctionnalités
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: overview
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: edad1223-a55f-4e4a-8e90-3471f2f559ac
description: 'L’élément Capabilities du schéma XML du fournisseur (OSC) permet à un fournisseur OSC de spécifier ses fonctionnalités. Ces fonctionnalités incluent les fonctionnalités suivantes :'
ms.openlocfilehash: ff6abbd4d99eb542a11e3d64a2015fc0585fca79
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33435004"
---
# <a name="xml-for-capabilities"></a>XML pour les fonctionnalités

**L’élément Capabilities** du schéma XML du fournisseur (OSC) permet à un fournisseur OSC de spécifier ses fonctionnalités. Ces fonctionnalités incluent les fonctionnalités suivantes : 
  
- Si le fournisseur prend en charge l’obtention, la mise en cache ou la recherche dynamique d’amis et d’activités à partir du réseau social.
    
- La façon dont OSC doit afficher certaines interfaces utilisateur d’accès.
    
- Si l’OSC doit utiliser l’authentification basée sur les formulaires ou configurer automatiquement le réseau social et les journaux sur l’utilisateur sur le réseau social.
    
Le schéma XML des **fonctionnalités** est essentiel car il identifie à l’OSC les fonctionnalités pris en charge par le fournisseur. Un fournisseur OSC doit implémenter la méthode [ISocialProvider::GetCapabilities](isocialprovider-getcapabilities.md) qui renvoie une  _chaîne de_ résultats. L’OSC appelle **ISocialProvider::GetCapabilities** pour obtenir des informations sur  les fonctionnalités du fournisseur OSC dans la chaîne de résultats, qui est conforme à la définition de schéma XML pour l’élément **capabilities.** Ces informations permettent aux appels ultérieurs de l’OSC au fournisseur OSC de fonctionner correctement. 
  
Pour spécifier les fonctionnalités d’un fournisseur OSC en tant que paramètre de sortie de la méthode **ISocialProvider::GetCapabilities,** vous devez vous conformer au schéma XML d’extensibilité du fournisseur OSC. La figure suivante illustre **la** structure XML des fonctionnalités. 
  
**Figure 1. \<fonctionnalités \> Structure XML**

![Structure XML des fonctionnalités](media/ol14oscref_Specifyingxmlforcapabilities_image1.gif)
  
Pour obtenir des descriptions détaillées des éléments enfants de l’élément **capabilities,** voir [Capabilities XML Elements](capabilities-xml-elements.md). Pour obtenir un exemple de **fonctionnalités** XML, voir [Capabilities XML Example](capabilities-xml-example.md). Pour obtenir une définition complète du schéma XML du fournisseur OSC, y compris les éléments requis ou facultatifs, voir Outlook [Social Connector Provider XML Schema](outlook-social-connector-provider-xml-schema.md).
  
## <a name="see-also"></a>Voir aussi

- [Exemple de fonctionnalités XML](capabilities-xml-example.md)  
- [Synchronisation des amis et des activités](synchronizing-friends-and-activities.md)  
- [XML pour les amis](xml-for-friends.md)  
- [XML pour les activités](xml-for-activities.md)
- [Développement d'un fournisseur avec le schéma XML OSC](developing-a-provider-with-the-osc-xml-schema.md)

