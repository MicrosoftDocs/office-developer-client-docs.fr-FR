---
title: Recordset2.Cancel, méthode (DAO)
TOCTitle: Cancel Method
ms:assetid: cae49f36-3aad-80d8-c15f-a7a584aa2e9b
ms:mtpsurl: https://msdn.microsoft.com/library/Ff834366(v=office.15)
ms:contentKeyID: 48547703
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: d203d1f1888539a4907da246e20ed711e61ee51f
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32307412"
---
# <a name="recordset2cancel-method-dao"></a>Recordset2.Cancel, méthode (DAO)


**S’applique à** : Access 2013, Office 2013

## <a name="syntax"></a>Syntaxe

*.* Annuler

*expression* Expression qui renvoie un **objet Recordset2.**

## <a name="remarks"></a>Remarques

Utilisez la **méthode Cancel** pour terminer l’exécution d’un appel asynchrone de méthode **Execute** ou **OpenConnection** (autrement dit, la méthode a été invoquée avec l’option dbRunAsync). **L’annulation** retourne une erreur d’utilisation si dbRunAsync n’a pas été utilisé dans la méthode que vous essayez de terminer.

