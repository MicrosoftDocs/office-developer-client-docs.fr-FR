---
title: 'Chapitre 7 : Gestion des événements ADO'
TOCTitle: 'Chapter 7: Handling ADO events'
ms:assetid: 22924fe2-d00d-8a0c-52f5-2dc6039537ff
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249004(v=office.15)
ms:contentKeyID: 48543709
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: e7357cc60a3bddbf96c2abae39fecfb7107025e2
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32296401"
---
# <a name="chapter-7-handling-ado-events"></a>Chapitre 7 : Gestion des événements ADO

**S’applique à** : Access 2013, Office 2013

Le modèle d’événements ADO prend en charge certaines opérations synchrones et asynchrones, qui génèrent des *événements*, ou notifications, avant et après leur exécution. Un événement représente, en fait, l’appel d’une routine de gestionnaire d’événements que vous définissez dans votre application.

Si vous fournissez des fonctions ou des procédures de gestionnaire pour les événements qui se produisent avant l'exécution de l'opération, vous pouvez examiner ou modifier les paramètres qui ont été transmis à celle-ci. En outre, étant donné que l'opération n'a pas encore été exécutée, vous pouvez l'annuler ou lui permettre d'aller à son terme.

Les événements qui se produisent au terme d'une opération sont particulièrement importants si vous utilisez ADO de façon asynchrone. Par exemple, une application qui lance une opération [Recordset.Open](open-method-ado-recordset.md) asynchrone est avertie de la fin de l'opération par un événement de fin d'exécution.

L'utilisation du modèle d'événements ADO, bien qu'elle constitue une surcharge pour votre application, offre plus de souplesse que les autres méthodes de traitement des opérations asynchrones, telles que le contrôle de la propriété [State](state-property-ado.md) d'un objet avec une boucle.

Ce chapitre présente les rubriques suivantes :

- [Résumé du gestionnaire d’événements ADO](ado-event-handler-summary.md)
- [Types d’événements](types-of-events.md)
- [Paramètres d’événement](event-parameters.md)
- [Collaboration des gestionnaires d’événements](how-event-handlers-work-together.md)
- [Insaisissation d’événement ADO par langue (ADO)](ado-event-instantiation-by-language-ado.md)
