---
title: Optimize, propriété - Exemple (VB)
TOCTitle: Optimize Property Example (VB)
ms:assetid: f4576247-6057-c1fe-013d-74feaab33174
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250240(v=office.15)
ms:contentKeyID: 48548686
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 8ace24040fa5820f43462819b4c45617d61b0a5b
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471086"
---
# <a name="optimize-property-example-vb"></a><span data-ttu-id="d6623-102">Optimize, propriété - Exemple (VB)</span><span class="sxs-lookup"><span data-stu-id="d6623-102">Optimize Property Example (VB)</span></span>


<span data-ttu-id="d6623-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="d6623-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="d6623-104">Cet exemple illustre la propriété dynamique Optimize des objets [Field](field-object-ado.md).</span><span class="sxs-lookup"><span data-stu-id="d6623-104">This example demonstrates the [Field](field-object-ado.md) objects dynamic Optimize property.</span></span> <span data-ttu-id="d6623-105">Le champ ***zip*** de la table ***Authors*** de la base de données ***Pubs*** n’est pas indexé.</span><span class="sxs-lookup"><span data-stu-id="d6623-105">The ***zip*** field of the ***Authors*** table in the ***Pubs*** database is not indexed.</span></span> <span data-ttu-id="d6623-106">Définition de la propriété [Optimize](optimize-property-dynamic-ado.md) sur **True** dans le champ ***zip*** autorise ADO pour créer un index qui améliore les performances de la méthode [Find](find-method-ado.md) .</span><span class="sxs-lookup"><span data-stu-id="d6623-106">Setting the [Optimize](optimize-property-dynamic-ado.md) property to **True** on the ***zip*** field authorizes ADO to build an index that improves the performance of the [Find](find-method-ado.md) method.</span></span>

```vb 
 
'BeginOptimizeVB 
Public Sub Main() 
 On Error GoTo ErrorHandler 
 
 'To integrate this code 
 'replace the data source and initial catalog values 
 'in the connection string 
 
 ' recordset and connection variables 
 Dim Cnxn As ADODB.Connection 
 Dim rstAuthors As ADODB.Recordset 
 Dim strCnxn As String 
 Dim strSQLAuthors As String 
 
 ' Open connection 
 strCnxn = "Provider='sqloledb';Data Source='MySqlServer';" & _ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';" 
 Set Cnxn = New ADODB.Connection 
 Cnxn.Open strCnxn 
 
 ' open recordset client-side to enable index creation 
 Set rstAuthors = New ADODB.Recordset 
 rstAuthors.CursorLocation = adUseClient 
 strSQLAuthors = "SELECT * FROM Authors" 
 rstAuthors.Open strSQLAuthors, Cnxn, adOpenStatic, adLockReadOnly, adCmdText 
 ' Create the index 
 rstAuthors!zip.Properties("Optimize") = True 
 ' Find Akiko Yokomoto 
 rstAuthors.Find "zip = '94595'" 
 
 ' show results 
 Debug.Print rstAuthors!au_fname & " " & rstAuthors!au_lname & " " & _ 
 rstAuthors!address & " " & rstAuthors!city & " " & rstAuthors!State 
 rstAuthors!zip.Properties("Optimize") = False 'Delete the index 
 
 ' clean up 
 rstAuthors.Close 
 Cnxn.Close 
 Set rstAuthors = Nothing 
 Set Cnxn = Nothing 
 Exit Sub 
 
ErrorHandler: 
 ' clean up 
 If Not rstAuthors Is Nothing Then 
 If rstAuthors.State = adStateOpen Then rstAuthors.Close 
 End If 
 Set rstAuthors = Nothing 
 
 If Not Cnxn Is Nothing Then 
 If Cnxn.State = adStateOpen Then Cnxn.Close 
 End If 
 Set Cnxn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
End Sub 
'EndOptimizeVB 
```
