---
title: ExecuteOptions, propriété (RDS)
TOCTitle: ExecuteOptions property (RDS)
ms:assetid: fb244cbd-9a03-9128-1373-694c9061c9da
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250285(v=office.15)
ms:contentKeyID: 48548864
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: 9cf773090ccb37bf4cad4aff41499ad01f966479
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32293251"
---
# <a name="executeoptions-property-rds"></a>ExecuteOptions, propriété (RDS)


**S’applique à** : Access 2013, Office 2013

Indique si l'exécution asynchrone est activée.

## <a name="settings-and-return-values"></a>Paramètres et valeurs de retour

Définit ou renvoie l'une des valeurs suivantes.

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p>Constante</p></th>
<th><p>Description</p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><strong>adcExecSync</strong></p></td>
<td><p>Exécute l’actualisation suivante du <a href="recordset-object-ado.md">Recordset</a> de façon asynchrone.</p></td>
</tr>
<tr class="even">
<td><p><strong>adcExecAsync</strong></p></td>
<td><p>Valeur par défaut. Exécute l'actualisation suivante du <strong>Recordset</strong> de façon asynchrone.</p></td>
</tr>
</tbody>
</table>

> [!NOTE]
> [!REMARQUE] Chaque fichier exécutable côté client utilisant ces constantes doit fournir les déclarations correspondantes. Vous pouvez couper et coller les déclarations de constante dont vous avez besoin dans le fichier Adcvbs.inc placé dans le dossier C:\Program Files\Common Files\System\MSADC.

## <a name="remarks"></a>Remarques

Si **ExecuteOptions** est défini sur **adcExecAsync**, l’appel suivant de **Refresh** est exécuté de façon asynchrone sur le **Recordset** de l’objet [RDS.DataControl](datacontrol-object-rds.md).

Si vous essayez d’appeler [Reset](reset-method-rds.md),[Refresh](refresh-method-rds.md), [SubmitChanges](submitchanges-method-rds.md), [CancelUpdate](cancelupdate-method-ado.md) ou [Recordset](recordset-sourcerecordset-properties-rds.md) pendant l’exécution d’une autre opération asynchrone susceptible de modifier le **Recordset** de l’objet [RDS.DataControl](datacontrol-object-rds.md), une erreur se produit.

S’il se produit une erreur pendant une opération asynchrone, la valeur [ReadyState](readystate-property-rds.md) de l’objet **RDS.DataControl** (**adcReadyStateLoaded**) est remplacée par **adcReadyStateComplete** et la valeur de la propriété **Recordset** reste la même : *Nothing*.

