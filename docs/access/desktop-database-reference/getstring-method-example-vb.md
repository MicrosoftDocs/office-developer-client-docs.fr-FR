---
title: GetString, méthode - Exemple (VB)
TOCTitle: GetString Method Example (VB)
ms:assetid: fa954e48-0810-9d71-4e24-f3ae2839105a
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250280(v=office.15)
ms:contentKeyID: 48548849
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 64fff0857754fc1a249a1d823a0c827e393d329a
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471754"
---
# <a name="getstring-method-example-vb"></a><span data-ttu-id="e3b56-102">GetString, méthode - Exemple (VB)</span><span class="sxs-lookup"><span data-stu-id="e3b56-102">GetString Method Example (VB)</span></span>


<span data-ttu-id="e3b56-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="e3b56-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="e3b56-104">Cet exemple illustre la méthode [GetString](getstring-method-ado.md).</span><span class="sxs-lookup"><span data-stu-id="e3b56-104">This example demonstrates the [GetString](getstring-method-ado.md) method.</span></span>

<span data-ttu-id="e3b56-105">Supposons que vous déboguez un problème d'accès aux données et que vous sounaitez un moyen rapide et simple d'imprimer le contenu actuel d'un objet [Recordset](recordset-object-ado.md) de petite taille.</span><span class="sxs-lookup"><span data-stu-id="e3b56-105">Assume you are debugging a data access problem and want a quick, simple way of printing the current contents of a small [Recordset](recordset-object-ado.md).</span></span>

```vb 
 
'BeginGetStringVB 
 
 'To integrate this code 
 'replace the data source and initial catalog values 
 'in the connection string 
 
Public Sub Main() 
 On Error GoTo ErrorHandler 
 
 ' connection variables 
 Dim Cnxn As ADODB.Connection 
 Dim rstAuthors As ADODB.Recordset 
 Dim strCnxn As String 
 Dim strSQLAuthors As String 
 Dim varOutput As Variant 
 
 ' specific variables 
 Dim strPrompt As String 
 Dim strState As String 
 
 ' open connection 
 Set Cnxn = New ADODB.Connection 
 strCnxn = "Provider='sqloledb';Data Source='MySqlServer';" & _ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';" 
 Cnxn.Open strCnxn 
 
 ' get user input 
 strPrompt = "Enter a state (CA, IN, KS, MD, MI, OR, TN, UT): " 
 strState = Trim(InputBox(strPrompt, "GetString Example")) 
 
 ' open recordset 
 Set rstAuthors = New ADODB.Recordset 
 strSQLAuthors = "SELECT au_fname, au_lname, address, city FROM Authors " & _ 
 "WHERE state = '" & strState & "'" 
 rstAuthors.Open strSQLAuthors, Cnxn, adOpenStatic, adLockReadOnly, adCmdText 
 
 If Not rstAuthors.EOF Then 
 ' Use all defaults: get all rows, TAB as column delimiter, 
 ' CARRIAGE RETURN as row delimiter, EMPTY-string as null delimiter 
 varOutput = rstAuthors.GetString(adClipString) 
 ' print output 
 Debug.Print "State = '" & strState & "'" 
 Debug.Print "Name Address City" & vbCr 
 Debug.Print varOutput 
 Else 
 Debug.Print "No rows found for state = '" & strState & "'" & vbCr 
 End If 
 
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
'EndGetStringVB 
```
