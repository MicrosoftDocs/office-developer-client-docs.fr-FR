---
title: State, propriété (ADO)
TOCTitle: State property (ADO)
ms:assetid: ade0a50c-e2d8-23ac-4ea9-b012fedcd5db
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249819(v=office.15)
ms:contentKeyID: 48547053
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- ado210.chm1231176
f1_categories:
- Office.Version=v15
localization_priority: Normal
ms.openlocfilehash: 0bce3f87a6530315a128396fe0e4de5390e0f47e
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32308532"
---
# <a name="state-property-ado"></a>State, propriété (ADO)


**S’applique à** : Access 2013, Office 2013

Indique, pour tous les objets applicables, si l'état de l'objet est ouvert ou fermé.

Indique, pour tous les objets applicables exécutant une méthode asynchrone, si l'état actuel de l'objet est « en cours de connexion », « en cours d'exécution » ou « en cours d'extraction ».

## <a name="return-value"></a>Valeur renvoyée

Renvoie une valeur de type **Long** pouvant être une valeur [ObjectStateEnum](objectstateenum.md). La valeur par défaut est **adStateClosed**.

## <a name="remarks"></a>Remarques

Vous pouvez utiliser à tout moment la propriété **State** pour déterminer l'état actuel d'un objet donné.

La propriété **State** de l'objet peut présenter plusieurs valeurs associées. Par exemple, si une instruction est en cours d'exécution, cette propriété aura la valeur **adStateOpen** associée à **adStateExecuting**.

La propriété **State** est en lecture seule.

