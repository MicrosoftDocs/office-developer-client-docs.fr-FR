---
title: GetRows, méthode - Exemple (JScript)
TOCTitle: GetRows Method Example (JScript)
ms:assetid: 72d7e2d9-1e19-e993-0b0e-5310405c9b75
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249466(v=office.15)
ms:contentKeyID: 48545620
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 09c8a463f049e168a44a3b2175c40000b59a4c72
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470461"
---
# <a name="getrows-method-example-jscript"></a><span data-ttu-id="5d498-102">GetRows, méthode - Exemple (JScript)</span><span class="sxs-lookup"><span data-stu-id="5d498-102">GetRows Method Example (JScript)</span></span>


<span data-ttu-id="5d498-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="5d498-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="5d498-104">Cet exemple utilise la méthode [GetRows](getrows-method-ado.md) pour récupérer toutes les lignes de la table *Customers* à partir d’un [objet Recordset](recordset-object-ado.md) et remplir un tableau avec les données résultantes.</span><span class="sxs-lookup"><span data-stu-id="5d498-104">This example uses the [GetRows](getrows-method-ado.md) method to retrieve all rows of the *Custiomers* table from a [Recordset](recordset-object-ado.md) and to fill an array with the resulting data.</span></span> <span data-ttu-id="5d498-105">La méthode **GetRows** renverra moins de lignes que le nombre voulu dans deux cas : si la [fin de fichier](bof-eof-properties-ado.md) a été atteinte ou si la méthode **GetRows** a tenté de récupérer un enregistrement qui a été supprimé par un autre utilisateur.</span><span class="sxs-lookup"><span data-stu-id="5d498-105">The **GetRows** method will return fewer than the desired number of rows in two cases: either if [EOF](bof-eof-properties-ado.md) has been reached, or if **GetRows** tried to retrieve a record that was deleted by another user.</span></span> <span data-ttu-id="5d498-106">La fonction retourne **False** uniquement dans le second cas.</span><span class="sxs-lookup"><span data-stu-id="5d498-106">The function returns **False** only if the second case occurs.</span></span> <span data-ttu-id="5d498-107">Coupez et collez le code suivant dans le Bloc-notes ou un autre éditeur de texte et enregistrez-le sous **GetRowsJS.asp**.</span><span class="sxs-lookup"><span data-stu-id="5d498-107">Cut and paste the following code to Notepad or another text editor, and save it as **GetRowsJS.asp**.</span></span>

```javascript 
 
<!-- BeginGetRowsJS --> 
<%@ LANGUAGE="JScript" %> 
<%// use this meta tag instead of adojavas.inc%> 
<!--METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" --> 
 
<html> 
 
<head> 
<title>ADO Recordset.GetRows Example (JScript)</title> 
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
 
<h1>ADO Recordset.GetRows Example (JScript)</h1> 
 <!-- Page text goes here --> 
<% 
 var Connect = "Provider='sqloledb';Data Source=" + Request.ServerVariables("SERVER_NAME") + ";" + 
 "Initial Catalog='Northwind';Integrated Security='SSPI';"; 
 var mySQL = "select * from customers;"; 
 var showblank = " "; 
 var shownull = "-null-"; 
 
 var connTemp = Server.CreateObject("ADODB.Connection"); 
 
 try 
 { 
 connTemp.Open(Connect); 
 var rsTemp = Server.CreateObject("ADODB.Recordset"); 
 rsTemp.ActiveConnection = connTemp; 
 rsTemp.CursorLocation = adUseClient; 
 rsTemp.CursorType = adOpenKeyset; 
 rsTemp.LockType = adLockOptimistic; 
 rsTemp.Open(mySQL); 
 
 rsTemp.MoveFirst(); 
 
 if (rsTemp.RecordCount == 0) 
 { 
 Response.Write("No records matched "); 
 Response.Write (mySQL & "So cannot make table..."); 
 connTemp.Close(); 
 Response.End(); 
 } else 
 { 
 Response.Write('<table width="100%" border="2">'); 
 Response.Write('<tr class="thead2">'); 
 
 // Headings On The Table for each Field Name 
 for (var i=0; i<rsTemp.Fields.Count; i++) 
 { 
 fieldObject = rsTemp.fields(i); 
 Response.Write('<td width="' + Math.floor(100 / rsTemp.Fields.Count) + '%">' + fieldObject.name + "</td>"); 
 } 
 
 Response.Write("</tr>"); 
 
 // JScript doesn't support multi-dimensional arrays 
 // so we'll convert the returned array to a single 
 // dimensional JScript array and then display the data. 
 tempArray = rsTemp.GetRows(); 
 recArray = tempArray.toArray(); 
 
 var col = 1; 
 var maxCols = rsTemp.Fields.Count; 
 
 for (var thisField=0; thisField<recArray.length; thisField++) 
 { 
 if (col == 1) 
 Response.Write('<tr class="tbody">'); 
 if (recArray[thisField] == null) 
 recArray[thisField] = shownull; 
 if (recArray[thisField] == "") 
 recArray[thisField] = showblank; 
 Response.Write("<td>" + recArray[thisField] + "</td>"); 
 col++ 
 if (col > maxCols) 
 { 
 Response.Write("</tr>"); 
 col = 1; 
 } 
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
 if (rsTemp.State == adStateOpen) 
 rsTemp.Close; 
 if (connTemp.State == adStateOpen) 
 connTemp.Close; 
 rsTemp = null; 
 connTemp = null; 
 } 
%> 
 
</body> 
 
</html> 
<!-- EndGetRowsJS --> 
 
```
