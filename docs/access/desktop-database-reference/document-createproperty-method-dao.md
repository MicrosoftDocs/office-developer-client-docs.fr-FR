---
title: Document.CreateProperty method (DAO)
TOCTitle: CreateProperty Method
ms:assetid: 834fda60-1edf-38df-a9a5-d9d15e55e425
ms:mtpsurl: https://msdn.microsoft.com/library/Ff196739(v=office.15)
ms:contentKeyID: 48546003
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052967
f1_categories:
- Office.Version=v15
localization_priority: Normal
ms.openlocfilehash: d1e2e2200953580406dded7d5c014b533129b133
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32293818"
---
# <a name="documentcreateproperty-method-dao"></a>Document.CreateProperty method (DAO)

**S’applique à** : Access 2013, Office 2013

Crée un objet utilisateur **[Property](property-object-dao.md)** (espaces de travail Microsoft Access uniquement).

## <a name="syntax"></a>Syntaxe

*.* CreateProperty(***Name***, ***Type***, ***Value***, ***DDL***)

*expression* Variable qui représente un **objet Document.**

## <a name="parameters"></a>Parameters

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th><p>Nom</p></th>
<th><p>Obligatoire/facultatif</p></th>
<th><p>Type de données</p></th>
<th><p>Description</p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><em>Name</em></p></td>
<td><p>Facultatif</p></td>
<td><p><strong>Variant</strong></p></td>
<td><p>Valeur de type <strong>String</strong> qui nomme de façon unique le nouvel objet <strong>Property</strong>. Voir la propriété <strong>Name</strong> pour plus d'informations sur les noms <strong>Property</strong> valides.</p></td>
</tr>
<tr class="even">
<td><p><em>Type</em></p></td>
<td><p>Facultatif</p></td>
<td><p><strong>Variant</strong></p></td>
<td><p>Constante qui définit le type de données du nouvel objet <strong>Property</strong>. Pour connaître les types de données valides, reportez-vous à la propriété <strong><a href="field-type-property-dao.md">Type</a></strong>.</p></td>
</tr>
<tr class="odd">
<td><p><em>Value</em></p></td>
<td><p>Facultatif</p></td>
<td><p><strong>Variant</strong></p></td>
<td><p><strong>Variant</strong> contenant la valeur initiale de la propriété. Pour plus d’informations, reportez-vous à la méthode <strong><a href="field-value-property-dao.md">Value</a></strong>.</p></td>
</tr>
<tr class="even">
<td><p><em>DDL</em></p></td>
<td><p>Facultatif</p></td>
<td><p><strong>Variant</strong></p></td>
<td><p>Valeur de type <strong>Variant</strong> (sous-type <strong>Boolean</strong>) indiquant si l'objet <strong>Property</strong> est ou non un objet DDL. La valeur par défaut est <strong>False</strong>. Si DDL a la valeur <strong>True</strong>, les utilisateurs ne peuvent pas modifier ou supprimer cet objet <strong>Property</strong> sauf s'ils détiennent une autorisation <strong>dbSecWriteDef</strong>.</p></td>
</tr>
</tbody>
</table>


## <a name="return-value"></a>Valeur renvoyée

Propriété

## <a name="remarks"></a>Remarques

Vous pouvez créer un objet utilisateur **Property** uniquement dans la collection **[Properties](properties-collection-dao.md)** d'un objet qui est persistant.

Si vous omettez une ou plusieurs parties facultatives lorsque vous utilisez **CreateProperty**, vous pouvez utiliser une instruction d'affectation appropriée pour définir ou réinitialiser la propriété correspondante avant d'ajouter le nouvel objet à une collection. Une fois que vous ajoutez l'objet, vous pouvez modifier certains de ses paramètres de propriété mais pas tous. Pour plus d'informations, reportez-vous aux rubriques des propriétés **Name**, **Type** et **Value**.

Si le nom fait référence à un objet qui est déjà membre de la collection, une erreur d’utilisation se produit lorsque vous utilisez **[la méthode Append.](fields-append-method-dao.md)**

Pour supprimer un objet **Property** défini par l’utilisateur de la collection, utilisez la méthode **[Delete](fields-delete-method-dao.md)** sur la collection **Properties.** Vous ne pouvez pas supprimer de propriétés intégrées.


> [!NOTE]
> Si vous omettez l’argument DDL, sa valeur par défaut est False (non DDL). Dans la mesure où la propriété DDL correspondante est exposée, vous devez supprimer et recréer l'objet **Property** pour passer de DDL à non DDL.


