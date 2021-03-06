---
title: Parse Function (Access custom web app)
manager: kelbow
ms.date: 09/05/2017
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: 09dee0ae-89b2-449c-a3c8-d6b270710b64
description: Elle pare une valeur de texte et renvoie sa valeur dans un type donné à l’aide de la culture de l’application.
ms.openlocfilehash: d664985ab1d7a7d33b99c52d5bab4aa714767e40
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33411133"
---
# <a name="parse-function-access-custom-web-app"></a>Parse Function (Access custom web app)

Elle pare une valeur de texte et renvoie sa valeur dans un type donné à l’aide de la culture de l’application.
  
> [!NOTE]
> La fonctionnalité de stockage cloud décrite dans cet article n’est plus prise en charge dans Office 2013 et Office 2016 et peut entraîner le message d’erreur suivant : > *Sorry, we’re having server problems, so we can’t add \<service\> right now. Please try again later.* (« Nous rencontrons actuellement des problèmes de serveur et nous sommes dans l’incapacité d’ajouter le service. Merci de réessayer ultérieurement. ») > En ce qui concerne le stockage cloud pour Office Online, Office pour iOS et Office pour Android, consultez notre [programme de partenariat de stockage cloud Office](https://dev.office.com/programs/officecloudstorage). 
  
## <a name="syntax"></a>Syntaxe

 **Parse** (*TextExpression*, *DataType*) 
  
La **fonction Parse** contient les arguments suivants. 
  
|**Nom de l’argument**|**Description**|
|:-----|:-----|
| *TextExpression*  <br/> |Expression de texte représentant la valeur mise en forme à analyser dans le type de données spécifié.  <br/> |
| *DataType*  <br/> |Valeur littérale représentant le type de données demandé pour le résultat.  <br/> |
   
## <a name="remarks"></a>Remarques

Utilisez **l’utilisation d’Parse** uniquement pour la conversion de chaîne en types de date/heure et de nombre. Pour les conversions de type général, utilisez **la fonction Convert.** Gardez à l’esprit qu’il existe une certaine surcharge de performances lors de l’exécution de l’une des valeurs de chaîne. 
  

