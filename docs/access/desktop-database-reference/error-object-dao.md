---
title: Objet Error - accès aux données (DAO) des objets
TOCTitle: Error Object
ms:assetid: e2608bc9-bece-9b47-4562-7a2689601f75
ms:mtpsurl: https://msdn.microsoft.com/library/Ff835711(v=office.15)
ms:contentKeyID: 48548289
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 62b195907d5acc05832c1feac45165aadd9e14d1
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470399"
---
# <a name="error-object-dao"></a><span data-ttu-id="dfbe4-102">Error Object (DAO)</span><span class="sxs-lookup"><span data-stu-id="dfbe4-102">Error Object (DAO)</span></span>


<span data-ttu-id="dfbe4-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="dfbe4-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="dfbe4-104">L'objet **Error** contient des détails sur les erreurs d'accès aux données, chacune d'entre elle appartenant à une opération simple impliquant DAO.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-104">**Error** object contains details about data access errors, each of which pertains to a single operation involving DAO.</span></span>

## <a name="remarks"></a><span data-ttu-id="dfbe4-105">Remarques</span><span class="sxs-lookup"><span data-stu-id="dfbe4-105">Remarks</span></span>

<span data-ttu-id="dfbe4-p101">Toute opération impliquant DAO peut générer une ou plusieurs erreurs. Par exemple, un appel à un serveur ODBC peut entraîner une erreur au niveau du serveur de base de données, d'ODBC, et de DAO. Lorsque chacun de ces types d'erreur se produit, un objet **Error** est placé dans la collection **Errors** de l'objet **DBEngine**. Un événement simple peut alors entraîner l'apparition de plusieurs objets **Error** dans la collection **Errors**.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-p101">Any operation involving DAO can generate one or more errors. For example, a call to an ODBC server might result in an error from the database server, an error from ODBC, and a DAO error. As each such error occurs, an **Error** object is placed in the **Errors** collection of the **DBEngine** object. A single event can therefore result in several **Error** objects appearing in the **Errors** collection.</span></span>

<span data-ttu-id="dfbe4-p102">Lorsqu'une opération DAO suivante génère une erreur, la collection **Errors** est effacée, et un ou plusieurs nouveaux objets **Error** sont placés dans la collection **Errors**. Les opérations DAO qui ne génèrent pas d'erreur n'ont aucun effet sur la collection **Errors**.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-p102">When a subsequent DAO operation generates an error, the **Errors** collection is cleared, and one or more new **Error** objects are placed in the **Errors** collection. DAO operations that don't generate an error have no effect on the **Errors** collection.</span></span>

<span data-ttu-id="dfbe4-p103">L'ensemble d'objets **Error** dans la collection **Errors** décrit une erreur. Le premier objet **Error** correspond à l'erreur de niveau inférieur (erreur d'origine), le second objet correspond à l'erreur de niveau immédiatement supérieur, etc. Par exemple, si une erreur ODBC se produit lorsque vous tentez d'ouvrir un objet **[Recordset](recordset-object-dao.md)**, le premier objet **Error**, **Errors**(0), contient l'erreur ODBC de niveau inférieur. Les erreurs suivantes contiennent les erreurs ODBC renvoyées par les différentes couches d'ODBC. Dans ce cas, le gestionnaire de pilotes ODBC, et probablement le pilote proprement dit, renvoient des objets **Error** distincts. Le dernier objet **Error**, **Errors.Count-** 1, contient l'erreur DAO qui indique que l'objet n'a pas pu être ouvert.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-p103">The set of **Error** objects in the **Errors** collection describes one error. The first **Error** object is the lowest level error (the originating error), the second the next higher level error, and so forth. For example, if an ODBC error occurs while trying to open a **[Recordset](recordset-object-dao.md)** object, the first **Error** object— **Errors**(0)— contains the lowest level ODBC error; subsequent errors contain the ODBC errors returned by the various layers of ODBC. In this case, the ODBC driver manager, and possibly the driver itself, return separate **Error** objects. The last **Error** object— **Errors.Count-** 1— contains the DAO error indicating that the object couldn't be opened.</span></span>

<span data-ttu-id="dfbe4-p104">L'énumération des erreurs spécifiques dans la collection **Errors** permet à vos programmes de traitement des erreurs de déterminer plus précisément la cause et l'origine d'une erreur, et entreprend la procédure appropriée pour effectuer la récupération. Vous pouvez lire les propriétés de l'objet **Error** afin d'obtenir des détails spécifiques sur chaque erreur, notamment :</span><span class="sxs-lookup"><span data-stu-id="dfbe4-p104">Enumerating the specific errors in the **Errors** collection enables your error-handling routines to more precisely determine the cause and origin of an error, and take appropriate steps to recover. You can read the **Error** object's properties to obtain specific details about each error, including:</span></span>

  - <span data-ttu-id="dfbe4-119">La propriété **[Description](error-description-property-dao.md)**, qui contient le texte de l'alerte d'erreur qui va être affichée à l'écran si l'erreur n'est pas interceptée.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-119">The **[Description](error-description-property-dao.md)** property, which contains the text of the error alert that will be displayed on the screen if the error is not trapped.</span></span>

  - <span data-ttu-id="dfbe4-120">La propriété **[Number](error-number-property-dao.md)**, qui contient la valeur d'entier long de la constante d'erreur.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-120">The **[Number](error-number-property-dao.md)** property, which contains the Long integer value of the error constant.</span></span>

  - <span data-ttu-id="dfbe4-p105">La propriété **[Source](error-source-property-dao.md)**, qui identifie l'objet qui a généré l'erreur. Cela s'avère particulièrement utile lorsque la collection **Errors** contient plusoeurs objets **Error** suite à une demande auprès d'une source de données ODBC.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-p105">The **[Source](error-source-property-dao.md)** property, which identifies the object that raised the error. This is particularly useful when you have several **Error** objects in the **Errors** collection following a request to an ODBC data source.</span></span>

  - <span data-ttu-id="dfbe4-123">Les propriétés **HelpFile** et **HelpContext**, qui indiquent respectivement le fichier d'aide et la rubrique d'aide Microsoft Windows appropriés (le cas échéant) pour l'erreur.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-123">The **HelpFile** and **HelpContext** properties, which indicate the appropriate Microsoft Windows Help file and Help topic, respectively, (if any exist) for the error.</span></span>
    

    > [!NOTE]
    > <P><span data-ttu-id="dfbe4-124">[!REMARQUE] Lors de la programmation dans Microsoft Visual Basic pour Applications (VBA), si vous utilisez le mot clé <STRONG>New</STRONG> pour créer un objet qui entraîne consécutivement une erreur avant que cet objet soit ajouté à une collection, la collection <STRONG>Errors</STRONG> de l'objet <STRONG>DBEngine</STRONG> ne contient pas d'entrée pour l'erreur de cet objet, car le nouvel objet est associé à l'objet <STRONG>DBEngine</STRONG>.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-124">When programming in Microsoft Visual Basic for Applications (VBA), if you use the <STRONG>New</STRONG> keyword to create an object that subsequently causes an error before that object has been appended to a collection, the <STRONG>DBEngine</STRONG> object's <STRONG>Errors</STRONG> collection won't contain an entry for that object's error, because the new object is not associated with the <STRONG>DBEngine</STRONG> object.</span></span> <span data-ttu-id="dfbe4-125">Cependant, les informations d'erreur sont disponibles dans l'objet VBA <STRONG>Err</STRONG>.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-125">However, the error information is available in the VBA <STRONG>Err</STRONG> object.</span></span> <span data-ttu-id="dfbe4-126">Le code de traitement des erreurs VBA doit examiner la collection <STRONG>Errors</STRONG> chaque fois que vous anticipez une erreur d'accès aux données.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-126">Your VBA error-handling code should examine the <STRONG>Errors</STRONG> collection whenever you anticipate a data access error.</span></span> <span data-ttu-id="dfbe4-127">Si vous écrivez un gestionnaire d'erreur centralisé, testez l'objet VBA <STRONG>Err</STRONG> pour déterminer si les informations d'erreur dans la collection <STRONG>Errors</STRONG> sont valides.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-127">If you are writing a centralized error handler, test the VBA <STRONG>Err</STRONG> object to determine if the error information in the <STRONG>Errors</STRONG> collection is valid.</span></span> <span data-ttu-id="dfbe4-128">Si la propriété <STRONG>Number</STRONG> du dernier élément de la collection <STRONG>Errors</STRONG> (DBEngine.Errors.Count - 1) et la valeur de l’objet <STRONG>Err</STRONG> correspondent, vous pouvez ensuite utiliser une série d’instructions <STRONG>Select Case</STRONG> pour identifier les erreurs DAO ou erreurs qui se sont produites.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-128">If the <STRONG>Number</STRONG> property of the last element of the <STRONG>Errors</STRONG> collection (DBEngine.Errors.Count - 1) and the value of the <STRONG>Err</STRONG> object match, you can then use a series of <STRONG>Select Case</STRONG> statements to identify the particular DAO error or errors that occurred.</span></span> <span data-ttu-id="dfbe4-129">Si elles ne correspondent pas, utilisez la méthode <STRONG><A href="errors-refresh-method-dao.md">Refresh</A></STRONG> dans la collection <STRONG>Errors</STRONG>.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-129">If they do not match, use the <STRONG><A href="errors-refresh-method-dao.md">Refresh</A></STRONG> method on the <STRONG>Errors</STRONG> collection.</span></span></P>



## <a name="example"></a><span data-ttu-id="dfbe4-130">Exemple</span><span class="sxs-lookup"><span data-stu-id="dfbe4-130">Example</span></span>

<span data-ttu-id="dfbe4-131">Cet exemple force une erreur, la capture et affiche les propriétés **Description**, **Number**, **Source**, **HelpContext** et **HelpFile** de l'objet **Error** qui en résulte.</span><span class="sxs-lookup"><span data-stu-id="dfbe4-131">This example forces an error, traps it, and displays the **Description**, **Number**, **Source**, **HelpContext**, and **HelpFile** properties of the resulting **Error** object.</span></span>

```vb 
Sub DescriptionX() 
 
   Dim dbsTest As Database 
 
   On Error GoTo ErrorHandler 
 
   ' Intentionally trigger an error. 
   Set dbsTest = OpenDatabase("NoDatabase") 
 
   Exit Sub 
 
ErrorHandler: 
   Dim strError As String 
   Dim errLoop As Error 
 
   ' Enumerate Errors collection and display properties of  
   ' each Error object. 
   For Each errLoop In Errors 
      With errLoop 
         strError = _ 
            "Error #" & .Number & vbCr 
         strError = strError & _ 
            "  " & .Description & vbCr 
         strError = strError & _ 
            "  (Source: " & .Source & ")" & vbCr 
         strError = strError & _ 
            "Press F1 to see topic " & .HelpContext & vbCr 
         strError = strError & _ 
            "  in the file " & .HelpFile & "." 
      End With 
      MsgBox strError 
   Next 
 
   Resume Next 
 
End Sub 
 
```
