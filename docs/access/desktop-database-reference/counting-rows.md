---
title: Comptage des lignes (référence de base de données de bureau Access)
TOCTitle: Counting rows
ms:assetid: ff684c5e-7f41-0dae-beea-f5c71f79bd84
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250312(v=office.15)
ms:contentKeyID: 48548963
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: 2388978185ac29149f7f15150ccfdbc559cc910f
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32295414"
---
# <a name="counting-rows"></a>Décompte des lignes


**S’applique à** : Access 2013, Office 2013

La propriété **RecordCount** renvoie une valeur de type **Long** qui indique le nombre d'enregistrements de l'objet **Recordset**. Utilisez la propriété **RecordCount** pour déterminer le nombre d'enregistrements d'un objet **Recordset**. La propriété renvoie -1 quand ADO ne peut pas déterminer le nombre d'enregistrements ou si le fournisseur ou le type de curseur ne prennent pas en charge la propriété **RecordCount**. Toute tentative de lecture de la propriété **RecordCount** sur un objet **Recordset** fermé génère une erreur.

La propriété **RecordCount** dépend des fonctionnalités du fournisseur et du type de curseur. La propriété **RecordCount** retourne -1 pour un curseur avant uniquement, le nombre réel pour un curseur statique ou de type jeu de clés et enfin, soit -1, soit le nombre réel pour un curseur dynamique, selon la source de données.

L’exemple d’objet **Recordset** présenté dans la section [Examen des données](chapter-3-examining-data.md) retourne –1 dans la mesure où un curseur avant uniquement a été ouvert. Pour pouvoir utiliser la propriété **RecordCount**, vous devez ouvrir l’objet **Recordset** avec un curseur plus élaboré (statique ou jeu de clés).

Dans certains cas, le fournisseur ou le curseur ne peut fournir la valeur **RecordCount** qu'en procédant au préalable à une extraction de tous les enregistrements à partir de la source de données. Pour forcer ce type d'extraction, il est recommandé d'appeler la méthode **MoveLast** sur l'objet **Recordset** avant d'appeler **RecordCount**.

Si vous remplacez la ligne de code qui appelle la méthode **Open** sur l'objet **Recordset** par les éléments suivants,

```vb 
 
oRs.Open sSQL, sCnStr, adOpenStatic, adLockOptimistic, adCmdText 
```

vous pouvez utiliser la propriété **RecordCount** dans la mesure où les curseurs statiques et le [fournisseur Microsoft OLE DB pour SQL Server](microsoft-ole-db-provider-for-sql-server.md) prennent en charge la propriété **RecordCount**. À titre d'exemple, le code suivant imprime le nombre d'enregistrements retourné par la commande dans la fenêtre de débogage, en supposant que le curseur prend en charge la propriété **RecordCount**:

```vb 
 
Debug.Print oRs.RecordCount ' Output: 4 
```

À ce stade, il est supposé que des paramètres de curseurs et verrous de ce type sont utilisés. En effet, ils offrent davantage de fonctionnalités même s'ils sont plus coûteux.

