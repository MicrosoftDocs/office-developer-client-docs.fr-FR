---
title: Ajout de données à un recordset
TOCTitle: Adding data to a Recordset
ms:assetid: a3d121a8-f52f-66cd-8849-c3a75aeb276a
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249761(v=office.15)
ms:contentKeyID: 48546797
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: 0d16eb5871bfe55af58a89cc06b413e8404df8cb
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32280215"
---
# <a name="adding-data-to-a-recordset"></a>Ajout de données à un recordset

**S’applique à** : Access 2013, Office 2013

Le **jeu d'enregistrements** est certainement l'objet ADO le plus utilisé. Dans ADO, un **jeu d'enregistrements** peut se définir comme une combinaison d'un jeu de résultats provenant d'une source de données et des comportements de curseurs associés. Vous pouvez ainsi mettre des données dans un **jeu d'enregistrements** et utiliser les méthodes et propriétés **Recordset** pour parcourir les lignes de données, afficher les valeurs de ces lignes ou encore manipuler le jeu de résultats.

Cette section s'intéresse plus particulièrement à l'ajout de données au **jeu d'enregistrements**. Pour en savoir plus sur la manière de parcourir ou mettre à jour les données, reportez-vous au [Chapitre 4 : Modification des données](chapter-4-editing-data.md) et au [Chapitre 5 : Mise à jour et conservation des données](chapter-5-updating-and-persisting-data.md). Les capacités avancées d'un objet **Command** ne sont pas toujours nécessaires pour ajouter votre jeu de résultats à un **jeu d'enregistrements**. Bien souvent, vous pouvez exécuter votre commande en définissant la propriété **Source** du **jeu d'enregistrements** ou en transférant une chaîne de commande à la méthode **Open** de l'objet **Recordset**.

Il existe différentes façons d'ajouter les données de votre source de données à un **jeu d'enregistrements**. La technique utilisée dépend des besoins de votre application et des capacités de votre fournisseur.

