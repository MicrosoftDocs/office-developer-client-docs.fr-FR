---
title: QueryDef.ODBCTimeout Property (DAO)
TOCTitle: ODBCTimeout Property
ms:assetid: b251c4fb-64a8-aa95-deed-64425df3e00c
ms:mtpsurl: https://msdn.microsoft.com/library/Ff822019(v=office.15)
ms:contentKeyID: 48547166
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1053052
f1_categories:
- Office.Version=v15
ms.openlocfilehash: dcb15511f3052366eb3d322c26cdd31d72c9beab
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472196"
---
# <a name="querydefodbctimeout-property-dao"></a><span data-ttu-id="bba3b-102">QueryDef.ODBCTimeout Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="bba3b-102">QueryDef.ODBCTimeout Property (DAO)</span></span>


<span data-ttu-id="bba3b-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="bba3b-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="bba3b-104">Indique le nombre de secondes d'attente avant que ne survienne une erreur d'expiration lors de l'exécution d'un objet **[QueryDef](querydef-object-dao.md)** sur une base de données ODBC.</span><span class="sxs-lookup"><span data-stu-id="bba3b-104">Indicates the number of seconds to wait before a timeout error occurs when a **[QueryDef](querydef-object-dao.md)** is executed on an ODBC database.</span></span>

## <a name="syntax"></a><span data-ttu-id="bba3b-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="bba3b-105">Syntax</span></span>

<span data-ttu-id="bba3b-106">*expression* . ODBCTimeout</span><span class="sxs-lookup"><span data-stu-id="bba3b-106">*expression* .ODBCTimeout</span></span>

<span data-ttu-id="bba3b-107">*expression* Variable qui représente un objet **QueryDef** .</span><span class="sxs-lookup"><span data-stu-id="bba3b-107">*expression* A variable that represents a **QueryDef** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="bba3b-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="bba3b-108">Remarks</span></span>

<span data-ttu-id="bba3b-p101">Lorsque la propriété **ODBCTimeout** est définie sur -1, le délai d'expiration revient par défaut au paramètre courant de la propriété **[QueryTimeout](database-querytimeout-property-dao.md)** de l'objet **[Connection](connection-object-dao.md)** ou **[Database](database-object-dao.md)** contenant l'objet **QueryDef**. Lorsque la propriété **ODBCTimeout** est définie sur 0, aucune erreur d'expiration ne survient.</span><span class="sxs-lookup"><span data-stu-id="bba3b-p101">When the **ODBCTimeout** property is set to -1, the timeout defaults to the current setting of the **[QueryTimeout](database-querytimeout-property-dao.md)** property of the **[Connection](connection-object-dao.md)** or **[Database](database-object-dao.md)** object that contains the **QueryDef**. When the **ODBCTimeout** property is set to 0, no timeout error occurs.</span></span>

<span data-ttu-id="bba3b-p102">Lorsque vous utilisez une base de données ODBC, comme Microsoft SQL Server, des retards peuvent survenir à cause d'un trafic réseau important ou d'une utilisation soutenue du serveur ODBC. Plutôt que d'attendre indéfiniment, vous pouvez définir le délai d'attente avant de renvoyer une erreur.</span><span class="sxs-lookup"><span data-stu-id="bba3b-p102">When you're using an ODBC database, such as Microsoft SQL Server, delays can occur because of network traffic or heavy use of the ODBC server. Rather than waiting indefinitely, you can specify how long to wait before returning an error.</span></span>

<span data-ttu-id="bba3b-113">La définition de la propriété **ODBCTimeout** d'un objet **QueryDef** écrase la valeur spécifiée par la propriété **QueryTimeout** de l'objet **Connection** ou **Database** contenant l'objet **QueryDef**, mais uniquement cet objet **QueryDef**.</span><span class="sxs-lookup"><span data-stu-id="bba3b-113">Setting the **ODBCTimeout** property of a **QueryDef** object overrides the value specified by the **QueryTimeout** property of the **Connection** or **Database** object containing the **QueryDef**, but only for that **QueryDef** object.</span></span>

## <a name="example"></a><span data-ttu-id="bba3b-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="bba3b-114">Example</span></span>

<span data-ttu-id="bba3b-115">L'exemple ci-dessous fait appel aux propriétés **ODBCTimeout** et **QueryTimeout** pour indiquer comment le paramètre **QueryTimeout** d'un objet **Database** définit le paramètre **ODBCTimeout** par défaut des objets **QueryDef** créés à partir de l'objet **Database**.</span><span class="sxs-lookup"><span data-stu-id="bba3b-115">This example uses the **ODBCTimeout** and **QueryTimeout** properties to show how the **QueryTimeout** setting on a **Database** object sets the default **ODBCTimeout** setting on any **QueryDef** objects created from the **Database** object.</span></span>

```vb 
Sub ODBCTimeoutX() 
 
 Dim dbsCurrent As Database 
 Dim qdfStores As QueryDef 
 Dim rstStores As Recordset 
 
 Set dbsCurrent = OpenDatabase("Northwind.mdb") 
 
 ' Change the default QueryTimeout of the Northwind 
 ' database. 
 Debug.Print "Default QueryTimeout of Database: " & _ 
 dbsCurrent.QueryTimeout 
 dbsCurrent.QueryTimeout = 30 
 Debug.Print "New QueryTimeout of Database: " & _ 
 dbsCurrent.QueryTimeout 
 
 ' Create a new QueryDef object. 
 Set qdfStores = dbsCurrent.CreateQueryDef("Stores", _ 
 "SELECT * FROM stores") 
 
 ' Note: The DSN referenced below must be configured to 
 ' use Microsoft Windows NT Authentication Mode to 
 ' authorize user access to the SQL Server. 
 qdfStores.Connect = _ 
 "ODBC;DATABASE=pubs;DSN=Publishers" 
 
 ' Change the ODBCTimeout setting of the new QueryDef 
 ' object from its default setting. 
 Debug.Print "Default ODBCTimeout of QueryDef: " & _ 
 qdfStores.ODBCTimeout 
 qdfStores.ODBCTimeout = 0 
 Debug.Print "New ODBCTimeout of QueryDef: " & _ 
 qdfStores.ODBCTimeout 
 
 ' Execute the query and display the results. 
 Set rstStores = qdfStores.OpenRecordset() 
 
 Debug.Print "Contents of recordset:" 
 With rstStores 
 Do While Not .EOF 
 Debug.Print , .Fields(0), .Fields(1) 
 .MoveNext 
 Loop 
 .Close 
 End With 
 
 ' Delete new QueryDef because this is a demonstration. 
 dbsCurrent.QueryDefs.Delete qdfStores.Name 
 dbsCurrent.Close 
 
End Sub 
 
```
