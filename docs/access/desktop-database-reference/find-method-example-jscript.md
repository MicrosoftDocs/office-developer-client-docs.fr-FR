---
title: Find, méthode – Exemple (JScript)
TOCTitle: Find method example (JScript)
ms:assetid: 87db96d6-4ed4-0807-8bff-62d978d4a008
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249589(v=office.15)
ms:contentKeyID: 48546116
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: e70a9bcc90348d6be48b7d681b6273b1fdd9201f
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32292390"
---
# <a name="find-method-example-jscript"></a>Find, méthode – Exemple (JScript)


**S’applique à** : Access 2013, Office 2013

Cet exemple utilise la méthode [Find](find-method-ado.md) de l’objet [Recordset](recordset-object-ado.md) pour localiser et afficher les sociétés de la base de données ***Northwind*** dont le nom commence par la lettre G. Coupez et collez le code suivant dans le Bloc-notes ou un autre éditeur de texte et enregistrez-le sous **FindJS.asp**.

```javascript 
 
<!-- BeginFindJS --> 
<%@ Language=JavaScript %> 
<%// use this meta tag instead of adojavas.inc%> 
<!--METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" --> 
 
<html> 
 
<head> 
<title>ADO Recordset.Find Example</title> 
<style> 
<!-- 
BODY { 
 font-family: 'Verdana','Arial','Helvetica',sans-serif; 
 BACKGROUND-COLOR:white; 
 COLOR:black; 
 } 
.thead { 
 background-color: #008080; 
 font-family: 'Verdana','Arial','Helvetica',sans-serif; 
 font-size: x-small; 
 color: white; 
 } 
.thead2 { 
 background-color: #800000; 
 font-family: 'Verdana','Arial','Helvetica',sans-serif; 
 font-size: x-small; 
 color: white; 
 } 
.tbody { 
 text-align: center; 
 background-color: #f7efde; 
 font-family: 'Verdana','Arial','Helvetica',sans-serif; 
 font-size: x-small; 
 } 
--> 
</style> 
</head> 
 
<body bgcolor="white"> 
 
<h1>ADO Recordset.Find Example</h1> 
<% 
 // connection and recordset variables 
 var Cnxn = Server.CreateObject("ADODB.Connection"); 
 var strCnxn = "Provider='sqloledb';Data Source=" + Request.ServerVariables("SERVER_NAME") + ";" + 
 "Initial Catalog='Northwind';Integrated Security='SSPI';"; 
 var rsCustomers = Server.CreateObject("ADODB.Recordset"); 
 // display string 
 var strMessage; 
 var strFind; 
 
 try 
 { 
 // open connection 
 Cnxn.Open(strCnxn); 
 
 //create recordset using object refs 
 SQLCustomers = "select * from Customers;"; 
 
 rsCustomers.ActiveConnection = Cnxn; 
 rsCustomers.CursorLocation = adUseClient; 
 rsCustomers.CursorType = adOpenKeyset; 
 rsCustomers.LockType = adLockOptimistic; 
 rsCustomers.Source = SQLCustomers; 
 
 rsCustomers.Open(); 
 rsCustomers.MoveFirst(); 
 
 //find criteria 
 strFind = "CompanyName like 'g%'" 
 rsCustomers.Find(strFind); 
 
 if (rsCustomers.EOF) { 
 Response.Write("No records matched "); 
 Response.Write(SQLCustomers & "So cannot make table..."); 
 Cnxn.Close(); 
 Response.End(); 
 } 
 else { 
 Response.Write('<table width="100%" border="2">'); 
 Response.Write('<tr class="thead2">'); 
 // Put Headings On The Table for each Field Name 
 for (thisField = 0; thisField < rsCustomers.Fields.Count; thisField++) { 
 fieldObject = rsCustomers.Fields(thisField); 
 Response.Write('<th width="' + Math.floor(100 / rsCustomers.Fields.Count) + '%">' + fieldObject.Name + "</th>"); 
 } 
 Response.Write("</tr>"); 
 
 while (!rsCustomers.EOF) { 
 Response.Write('<tr class="tbody">'); 
 for(thisField=0; thisField<rsCustomers.Fields.Count; thisField++) { 
 fieldObject = rsCustomers.Fields(thisField); 
 strField = fieldObject.Value; 
 if (strField == null) 
 strField = "-Null-"; 
 if (strField == "") 
 strField = ""; 
 Response.Write("<td>" + strField + "</td>"); 
 } 
 rsCustomers.Find(strFind, 1, adSearchForward) 
 Response.Write("</tr>"); 
 } 
 Response.Write("</table>"); 
 } 
 } 
 catch (e) 
 { 
 Response.Write(e.message); 
 } 
 finally 
 { 
 // clean up 
 if (rsCustomers.State == adStateOpen) 
 rsCustomers.Close; 
 if (Cnxn.State == adStateOpen) 
 Cnxn.Close; 
 rsCustomers = null; 
 Cnxn = null; 
 } 
%> 
 
</body> 
 
</html> 
<!-- EndFindJS --> 
 
```

