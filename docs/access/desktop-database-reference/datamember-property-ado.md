---
title: DataMember, propriété (ADO)
TOCTitle: DataMember property (ADO)
ms:assetid: f89e1d42-7993-764b-4e8a-2f449903f792
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250263(v=office.15)
ms:contentKeyID: 48548787
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: 410f11af8daf3912dca9dc78a1cb9216ff8f8dd1
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32294490"
---
# <a name="datamember-property-ado"></a>DataMember, propriété (ADO)

**S’applique à** : Access 2013, Office 2013

Indique le nom du membre de données qui sera extrait de l'objet référencé par la propriété [DataSource](datasource-property-ado.md).

## <a name="settings-and-return-values"></a>Paramètres et valeurs de retour

Définit ou renvoie une valeur de type **String**. Le nom ne respecte pas la casse.

## <a name="remarks"></a>Remarques

Cette propriété est utilisée pour créer des contrôles liés aux données avec l’environnement de données. L’environnement de données conserve des collections de données (sources de données) contenant des objets nommés (*membres de données*) qui seront représentés en tant qu’objet [Recordset](recordset-object-ado.md)*.*

Les propriétés **DataMember** et **DataSource** doivent être utilisées conjointement.

La propriété **DataMember** détermine l'objet spécifié par la propriété **DataSource** à représenter en tant qu'objet **Recordset**. L'objet **Recordset** doit être fermé avant que cette propriété soit définie. Une erreur est générée si la propriété **DataMember** n'est pas définie avant la propriété **DataSource** ou si le nom **DataMember** n'est pas reconnu par l'objet spécifié dans la propriété **DataSource**.

**Utilisation**

```vb
    Dim rs as New ADODB.Recordset
    rs.DataMember = "Command"     'Name of the rowset to bind to
    Set rs.DataSource = myDE      'Name of the object containing an IRowset
```
