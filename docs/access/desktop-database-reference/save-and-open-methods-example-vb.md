---
title: Save et Open, méthodes - Exemple (VB)
TOCTitle: Save and Open Methods Example (VB)
ms:assetid: aecb56b4-3ccd-d8f1-84a9-fd5a40aeca5f
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249828(v=office.15)
ms:contentKeyID: 48547081
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: d293cf67d328beb2e380e484a595eaa09bd6f970
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470113"
---
# <a name="save-and-open-methods-example-vb"></a><span data-ttu-id="802ca-102">Save et Open, méthodes - Exemple (VB)</span><span class="sxs-lookup"><span data-stu-id="802ca-102">Save and Open Methods Example (VB)</span></span>


<span data-ttu-id="802ca-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="802ca-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="802ca-104">Ces trois exemples montrent comment utiliser conjointement les méthodes [Save](save-method-ado.md) et [Open](open-method-ado-recordset.md).</span><span class="sxs-lookup"><span data-stu-id="802ca-104">These three examples demonstrate how the [Save](save-method-ado.md) and [Open](open-method-ado-recordset.md) methods can be used together.</span></span>

<span data-ttu-id="802ca-p101">Supposons que vous partez en voyage d'affaires et que vous souhaitez emporter une table de la base de données. Avant de partir, vous accédez aux données sous la forme d'un objet [Recordset](recordset-object-ado.md) et vous l'enregistrez dans un format transportable. Lorsque vous arrivez à destination, vous accédez à l'objet **Recordset** en tant qu'objet **Recordset** déconnecté, local. Vous modifiez l'objet **Recordset** avant de le réenregistrer. Enfin, un fois de retour, vous vous reconnectez à la base de données et vous la mettez à jour avec les modifications apportées au cours de votre voyage.</span><span class="sxs-lookup"><span data-stu-id="802ca-p101">Assume you are going on a business trip and want to take along a table from a database. Before you go, you access the data as a [Recordset](recordset-object-ado.md) and save it in a transportable form. When you arrive at your destination, you access the **Recordset** as a local, disconnected **Recordset**. You make changes to the **Recordset**, then save it again. Finally, when you return home, you connect to the database again and update it with the changes you made on the road.</span></span>

<span data-ttu-id="802ca-110">Pour commencer, accédez à la table ***Authors*** et enregistrez-la.</span><span class="sxs-lookup"><span data-stu-id="802ca-110">First, access and save the ***Authors*** table.</span></span>

```vb 
 
'BeginSaveVB 
 
 'To integrate this code 
 'replace the data source and initial catalog values 
 'in the connection string 
 
Public Sub Main() 
 On Error GoTo ErrorHandler 
 
 'recordset and connection variables 
 Dim rstAuthors As ADODB.Recordset 
 Dim Cnxn As ADODB.Connection 
 Dim strCnxn As String 
 Dim strSQLAuthors As String 
 
 ' Open connection 
 Set Cnxn = New ADODB.Connection 
 strCnxn = "Provider='sqloledb';Data Source='MySqlServer';" & _ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';" 
 Cnxn.Open strCnxn 
 
 Set rstAuthors = New ADODB.Recordset 
 strSQLAuthors = "SELECT au_id, au_lname, au_fname, city, phone FROM Authors" 
 rstAuthors.Open strSQLAuthors, Cnxn, adOpenDynamic, adLockOptimistic, adCmdText 
 
 'For sake of illustration, save the Recordset to a diskette in XML format 
 rstAuthors.Save "c:\Pubs.xml", adPersistXML 
 
 ' clean up 
 rstAuthors.Close 
 Cnxn.Close 
 Set rstAuthors = Nothing 
 Set Cnxn = Nothing 
 Exit Sub 
 
ErrorHandler: 
 'clean up 
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
'EndSaveVB 
```

<br/>

<span data-ttu-id="802ca-111">À ce stade, vous êtes arrivé à destination.</span><span class="sxs-lookup"><span data-stu-id="802ca-111">At this point, you have arrived at your destination.</span></span> <span data-ttu-id="802ca-112">Vous accédez à la table ***Authors*** **Recordset**déconnecté, local.</span><span class="sxs-lookup"><span data-stu-id="802ca-112">You will access the ***Authors*** table as a local, disconnected **Recordset**.</span></span> <span data-ttu-id="802ca-113">Le fournisseur **MSPersist** doit être installé sur l’ordinateur que vous utilisez pour accéder au fichier enregistré, r :\\Pubs.xml.</span><span class="sxs-lookup"><span data-stu-id="802ca-113">Don't forget you must have the **MSPersist** provider on the machine that you are using in order to access the saved file, a:\\Pubs.xml.</span></span>

```vb 
 
'BeginSave2VB 
Public Sub Main() 
 On Error GoTo ErrorHandler 
 
 Dim rst As ADODB.Recordset 
 Set rst = New ADODB.Recordset 
 
 'For sake of illustration, we specify all parameters 
 rst.Open "c:\Pubs.xml", "Provider=MSPersist;", adOpenForwardOnly, adLockBatchOptimistic, adCmdFile 
 
 'Now you have a local, disconnected Recordset - Edit as you desired 
 '(In this example the change makes no difference) 
 rst.Find "au_lname = 'Carson'" 
 If rst.EOF Then 
 Debug.Print "Name not found." 
 Exit Sub 
 End If 
 
 rst!city = "Chicago" 
 rst.Update 
 
 'Save changes in ADTG format this time, purely for sake of illustration. 
 'Note that the previous version is still on the diskette, as a:\Pubs.xml. 
 rst.Save "c:\Pubs.adtg", adPersistADTG 
 
 ' clean up 
 rst.Close 
 Set rst = Nothing 
 Exit Sub 
 
ErrorHandler: 
 'clean up 
 If Not rst Is Nothing Then 
 If rst.State = adStateOpen Then rst.Close 
 End If 
 Set rst = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
End Sub 
'EndSave2VB 
```

<br/>

<span data-ttu-id="802ca-p103">Enfin, vous êtes de retour et vous mettez à présent la base de données à jour avec vos modifications.</span><span class="sxs-lookup"><span data-stu-id="802ca-p103">Finally, you return home. Now update the database with your changes.</span></span>

```vb 
 
'BeginSave3VB 
Public Sub Main() 
 On Error GoTo ErrorHandler 
 
 'To integrate this code 
 'replace the data source and initial catalog values 
 'in the connection string 
 
 Dim Cnxn As New ADODB.Connection 
 Dim rst As ADODB.Recordset 
 Dim strCnxn As String 
 
 Set rst = New ADODB.Recordset 
 ' The lock mode is batch optimistic because we are going to 
 ' use the UpdateBatch method. 
 rst.Open "c:\Pubs.adtg", "Provider=MSPersist;", adOpenForwardOnly, adLockBatchOptimistic, adCmdFile 
 
 ' Connect to the database, associate the Recordset with the connection 
 ' then update the database table with the changed Recordset 
 strCnxn = "Provider=SQLOLEDB;Data Source=MySqlServer;Integrated Security=SSPI;Initial Catalog=pubs;" 
 Cnxn.Open strCnxn 
 
 rst.ActiveConnection = Cnxn 
 rst.UpdateBatch 
 
 ' clean up 
 rst.Close 
 Cnxn.Close 
 Set rst = Nothing 
 Set Cnxn = Nothing 
 Exit Sub 
 
ErrorHandler: 
 'clean up 
 If Not rst Is Nothing Then 
 If rst.State = adStateOpen Then rst.Close 
 End If 
 Set rst = Nothing 
 
 If Not Cnxn Is Nothing Then 
 If Cnxn.State = adStateOpen Then Cnxn.Close 
 End If 
 Set Cnxn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
End Sub 
'EndSave3VB 
```
