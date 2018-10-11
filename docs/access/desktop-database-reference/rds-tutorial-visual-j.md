---
title: Didacticiel RDS (Visual J++)
TOCTitle: RDS Tutorial (Visual J++)
ms:assetid: b5679bfe-e830-05df-8a1c-0744c96abe90
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249870(v=office.15)
ms:contentKeyID: 48547248
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 15eadc36079faa529585e539b67eb0da246cf4bc
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470674"
---
# <a name="rds-tutorial-visual-j"></a><span data-ttu-id="0d3f7-102">Didacticiel RDS (Visual J++)</span><span class="sxs-lookup"><span data-stu-id="0d3f7-102">RDS Tutorial (Visual J++)</span></span>


<span data-ttu-id="0d3f7-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="0d3f7-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="0d3f7-p101">ADO/WFC ne suit pas intégralement le modèle d'objet RDS en cela qu'il n'implémente pas l'objet [RDS.DataControl](datacontrol-object-rds.md). ADO/WFC implémente uniquement la classe côté client[RDS.DataSpace](dataspace-object-rds.md).</span><span class="sxs-lookup"><span data-stu-id="0d3f7-p101">ADO/WFC does not completely follow the RDS object model in that it does not implement the [RDS.DataControl](datacontrol-object-rds.md) object. ADO/WFC only implements the client-side class, [RDS.DataSpace](dataspace-object-rds.md).</span></span>

<span data-ttu-id="0d3f7-p102">La classe **DataSpace** implémente une seule méthode, [CreateObject](createobject-method-rds.md), qui retourne un objet [ObjectProxy](https://msdn.microsoft.com/library/jj249624\(v=office.15\)). La classe **DataSpace** implémente également la propriété [InternetTimeout](internettimeout-property-rds.md).</span><span class="sxs-lookup"><span data-stu-id="0d3f7-p102">The **DataSpace** class implements one method, [CreateObject](createobject-method-rds.md), which returns an [ObjectProxy](https://msdn.microsoft.com/library/jj249624\(v=office.15\)) object. The **DataSpace** class also implements the [InternetTimeout](internettimeout-property-rds.md) property.</span></span>

<span data-ttu-id="0d3f7-108">La classe **ObjectProxy** implémente une seule méthode, call, qui peut appeler n'importe quel objet métier côté serveur.</span><span class="sxs-lookup"><span data-stu-id="0d3f7-108">The **ObjectProxy** class implements one method, call, which can invoke any server-side business object.</span></span>

<span data-ttu-id="0d3f7-109">**Point de départ du didacticiel.**</span><span class="sxs-lookup"><span data-stu-id="0d3f7-109">**This is the beginning of the tutorial.**</span></span>

```java 
 
import com.ms.wfc.data.*; 
public class RDSTutorial 
{ 
 public void tutorial() 
 { 
// Step 1: Specify a server program. 
 ObjectProxy obj = 
 DataSpace.createObject( 
 "RDSServer.DataFactory", 
 "https://YourServer"); 
 
// Step 2: Server returns a Recordset. 
 Recordset rs = (Recordset) obj.call( 
 "Query", 
 new Object[] {"DSN=Pubs;", "SELECT * FROM Authors"}); 
 
// Step 3: Changes are sent to the server. 
 ... // Edit Recordset. 
 obj.call( 
 "SubmitChanges", 
 new Object[] {"DSN=Pubs;", rs}); 
 return; 
 } 
} 
```

<span data-ttu-id="0d3f7-110">**Fin du didacticiel.**</span><span class="sxs-lookup"><span data-stu-id="0d3f7-110">**This is the end of the tutorial.**</span></span>
