---
title: Relation.Table, propriété (DAO)
TOCTitle: Table Property
ms:assetid: cc4f64ef-c4e9-1a14-9263-5f8220d89840
ms:mtpsurl: https://msdn.microsoft.com/library/Ff834423(v=office.15)
ms:contentKeyID: 48547736
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1053068
f1_categories:
- Office.Version=v15
localization_priority: Normal
ms.openlocfilehash: 91a3262d92350618c2385013983020669b28ea5c
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32306992"
---
# <a name="relationtable-property-dao"></a>Relation.Table, propriété (DAO)


**S’applique à** : Access 2013, Office 2013

Indique le nom d'une table primaire d'un objet **[Relation](relation-object-dao.md)**. Ce nom doit correspondre à la valeur de la propriété **[Name](connection-name-property-dao.md)** d'un objet **[TableDef](tabledef-object-dao.md)** ou **[QueryDef](querydef-object-dao.md)** (espaces de travail Microsoft Access uniquement).

## <a name="syntax"></a>Syntaxe

*.* Table

*expression* Variable qui représente un objet **Relation.**

## <a name="remarks"></a>Remarques

La propriété **Table** est en lecture/écriture pour tout nouvel objet **Relation** non encore ajouté à une collection et en lecture seule pour tout objet **Relation** existant dans une collection **[Relations](relations-collection-dao.md)**.

Utilisez la propriété **Table** avec la propriété **[ForeignTable](relation-foreigntable-property-dao.md)** pour définir un objet **Relation**, qui représente la relation entre des champs dans deux tables ou requêtes. Donnez à la propriété **Table** la valeur de la propriété **Name** de l'objet primaire **TableDef** ou **QueryDef**, et attribuez à la propriété **ForeignTable** la valeur de la propriété **Name** de l'objet étranger (de référence) **TableDef** ou **QueryDef**. La propriété **[Attributes](field-attributes-property-dao.md)** détermine le type de relation entre les deux objets.

Par exemple, si vous avez une liste de codes de pièces valide (dans un champ appelé PartNo) stockée dans une table ValidParts, vous pouvez établir une relation un-à-plusieurs avec une table OrderItem de telle façon que si un code de pièce est entré dans la table OrderItem, il doit déjà se trouver dans la table ValidParts. Si le code de pièce n'existe pas dans la table ValidParts et que vous n'avez pas défini la propriété **Attributes** de l'objet **Relation** avec la valeur **dbRelationDontEnforce**, une erreur interceptable se produit.

Dans ce cas, la table ValidParts est la table primaire, de sorte que la propriété **Table** de l'objet **Relation** sont définie sur ValidParts et la propriété **ForeignTable** de l'objet **Relation** sur OrderItem. Les propriétés **Name** et **ForeignName** de l'objet **[Field](field-object-dao.md)** dans la collection [**Fields**](fields-collection-dao.md) de l'objet **Relation** sont définies sur PartNo.

## <a name="example"></a>Exemple

Cet exemple montre comment les propriétés **Table**, **ForeignTable** et **ForeignName** définissent les termes d’une **Relation** entre deux tables.

```vb
    Sub ForeignNameX() 
     
     Dim dbsNorthwind As Database 
     Dim relLoop As Relation 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     
     Debug.Print "Relation" 
     Debug.Print " Table - Field" 
     Debug.Print " Primary (One) "; 
     Debug.Print ".Table - .Fields(0).Name" 
     Debug.Print " Foreign (Many) "; 
     Debug.Print ".ForeignTable - .Fields(0).ForeignName" 
     
     ' Enumerate the Relations collection of the Northwind 
     ' database to report on the property values of 
     ' the Relation objects and their Field objects. 
     For Each relLoop In dbsNorthwind.Relations 
     With relLoop 
     Debug.Print 
     Debug.Print .Name & " Relation" 
     Debug.Print " Table - Field" 
     Debug.Print " Primary (One) "; 
     Debug.Print .Table & " - " & .Fields(0).Name 
     Debug.Print " Foreign (Many) "; 
     Debug.Print .ForeignTable & " - " & _ 
     .Fields(0).ForeignName 
     End With 
     Next relLoop 
     
     dbsNorthwind.Close 
     
    End Sub
```
