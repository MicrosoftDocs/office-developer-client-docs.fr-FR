---
title: Create, méthode – Exemple (VB)
TOCTitle: Create method example (VB)
ms:assetid: 3e6a4f3d-3b25-2dfb-5ef3-6a4c5326b78f
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249171(v=office.15)
ms:contentKeyID: 48544372
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: 325551763f8edded4fd6ed5c7191d827c339ddf4
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32295435"
---
# <a name="create-method-example-vb"></a>Create, méthode – Exemple (VB)


**S’applique à** : Access 2013, Office 2013

Le code suivant montre comment créer une nouvelle base de données Microsoft Jet à l'aide de la méthode [Create](create-method-adox.md).

```vb 
 
' BeginCreateDatabaseVB 
Sub CreateDatabase() 
 On Error GoTo CreateDatabaseError 
 
 Dim cat As New ADOX.Catalog 
 cat.Create "Provider='Microsoft.Jet.OLEDB.4.0';Data Source='c:\new.mdb'" 
 
 'Clean up 
 Set cat = Nothing 
 Exit Sub 
 
CreateDatabaseError: 
 Set cat = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
End Sub 
' EndCreateDatabaseVB 
```

