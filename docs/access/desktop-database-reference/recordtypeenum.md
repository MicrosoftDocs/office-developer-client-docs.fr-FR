---
title: RecordTypeEnum (référence de base de données de bureau Access)
TOCTitle: RecordTypeEnum
ms:assetid: 7edd6508-1507-4649-f1aa-03f1873ef09c
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249534(v=office.15)
ms:contentKeyID: 48545890
ms.date: 10/18/2018
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: c95e4b78a3e2259c3dc5961e87b98ca65ea55692
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32309282"
---
# <a name="recordtypeenum"></a>RecordTypeEnum

**S’applique à** : Access 2013, Office 2013

Spécifie le type de l’objet [Record](record-object-ado.md).

<br/>

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header">
<th><p>Constante</p></th>
<th><p>Valeur</p></th>
<th><p>Description</p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><strong>adSimpleRecord</strong></p></td>
<td><p>0</p></td>
<td><p>Indique un enregistrement <em>simple</em> (sans nœud enfant).</p></td>
</tr>
<tr class="even">
<td><p><strong>adCollectionRecord</strong></p></td>
<td><p>1</p></td>
<td><p>Indique un enregistrement à <em>collection</em> (contient des nœuds enfants).</p></td>
</tr>
<tr class="odd">
<td><p><strong>adRecordUnknown</strong></p></td>
<td><p>-1</p></td>
<td><p>Indique que le type de cet objet <strong>Record</strong> est inconnu.</p></td>
</tr>
<tr class="even">
<td><p><strong>adStructDoc</strong></p></td>
<td><p>2</p></td>
<td><p>Indique un enregistrement spécial à <em>collection</em>, qui représente des documents COM structurés.</p></td>
</tr>
</tbody>
</table>


### <a name="adowfc-equivalent"></a>Équivalent ADO/WFC

Ces constantes ne possèdent pas d'équivalent ADO/WFC.

