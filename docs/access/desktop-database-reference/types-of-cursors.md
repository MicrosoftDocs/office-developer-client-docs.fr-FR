---
title: Types de curseurs (référence de base de données de bureau Access)
TOCTitle: Types of Cursors
ms:assetid: 589f3755-3cf5-9470-bd66-8e8afa218fc5
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249299(v=office.15)
ms:contentKeyID: 48544996
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: 6493d746f43ac8d923e5b1b9d6dcd3de44b411ec
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32314027"
---
# <a name="types-of-cursors"></a>Types de curseurs


**S’applique à** : Access 2013, Office 2013

En règle générale, votre application doit utiliser le curseur le plus simple possible, capable de fournir l'accès aux données requis. Toutes les caractéristiques supplémentaires au-delà des fonctions de base (avant uniquement, lecture seule, statique, défilement, sans tampon) sont susceptibles d'avoir une incidence sur la mémoire du client, la charge réseau ou les performances. Dans la plupart des cas, les options du curseur par défaut génèrent un curseur plus complexe que celui nécessaire à votre application.

Le choix du type de curseur dépend de plusieurs éléments : la façon dont votre application utilise le jeu de résultats et diverses considérations de conception, notamment la taille du jeu de résultats, le pourcentage de données susceptibles d'être utilisées, le degré de sensibilité aux modifications de données et les performances d'application requises.

Au niveau le plus élémentaire, le choix de curseur varie selon que vous devez modifier les données ou simplement les consulter :

  - Si vous devez parcourir un jeu de résultats sans modifier de données, utilisez un curseur [avant uniquement](forward-only-cursors.md) ou un curseur [statique](static-cursors.md).

  - Si votre jeu de résultats est volumineux et que vous devez sélectionner uniquement quelques lignes, utilisez un curseur de type [jeu de clés](keyset-cursors.md).

  - Si vous souhaitez synchroniser un jeu de résultats avec les ajouts, les modifications et les suppressions récents effectués simultanément par tous les utilisateurs, utilisez un curseur [dynamique](dynamic-cursors.md).

Bien que chaque curseur semble différent, rappelez-vous que ces types de curseur ne présentent pas de différences majeures et possèdent souvent des caractéristiques et des options communes.

