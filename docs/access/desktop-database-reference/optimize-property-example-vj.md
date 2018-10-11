---
title: Optimize, propriété - Exemple (VJ++)
TOCTitle: Optimize Property Example (VJ++)
ms:assetid: d4ac9ae3-3304-addf-0292-7af4ed4fdbc2
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250067(v=office.15)
ms:contentKeyID: 48547949
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 67b587fb7a0e4a87bafab83ba21cebc6c33a16ec
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470455"
---
# <a name="optimize-property-example-vj"></a><span data-ttu-id="d6073-102">Optimize, propriété - Exemple (VJ++)</span><span class="sxs-lookup"><span data-stu-id="d6073-102">Optimize Property Example (VJ++)</span></span>


<span data-ttu-id="d6073-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="d6073-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="d6073-104">Cet exemple illustre la propriété dynamique Optimize des objets [Field](field-object-ado.md).</span><span class="sxs-lookup"><span data-stu-id="d6073-104">This example demonstrates the [Field](field-object-ado.md) object dynamic Optimize property.</span></span> <span data-ttu-id="d6073-105">Le champ ***zip*** de la table ***Authors*** de la base de données ***Pubs*** n’est pas indexé.</span><span class="sxs-lookup"><span data-stu-id="d6073-105">The ***zip*** field of the ***Authors*** table in the ***Pubs*** database is not indexed.</span></span> <span data-ttu-id="d6073-106">Définition de la propriété [Optimize](optimize-property-dynamic-ado.md) sur **True** dans le champ ***zip*** autorise ADO pour créer un index qui améliore les performances de la méthode [Find](find-method-ado.md) .</span><span class="sxs-lookup"><span data-stu-id="d6073-106">Setting the [Optimize](optimize-property-dynamic-ado.md) property to **True** on the ***zip*** field authorizes ADO to build an index that improves the performance of the [Find](find-method-ado.md) method.</span></span>

```java 
 
// BeginOptimizeJ 
import com.ms.wfc.data.*; 
import java.io.* ; 
 
public class OptimizeX 
{ 
 // The main entry point for the application. 
 
 public static void main (String[] args) 
 { 
 OptimizeX(); 
 System.exit(0); 
 } 
 
 // OptimizeX function 
 
 static void OptimizeX() 
 { 
 
 // Define ADO Objects. 
 Recordset rstAuthors = null; 
 
 // Declarations. 
 BufferedReader in = 
 new BufferedReader (new InputStreamReader(System.in)); 
 String strCnn = "Provider='sqloledb';Data Source='MySqlServer';" 
 + "Initial Catalog='Pubs';Integrated Security='SSPI';"; 
 
 try 
 { 
 rstAuthors = new Recordset(); 
 rstAuthors.setCursorLocation(AdoEnums.CursorLocation.CLIENT); 
 // Enable index creation. 
 rstAuthors.open("SELECT * FROM Authors", 
 strCnn, 
 AdoEnums.CursorType.STATIC, 
 AdoEnums.LockType.READONLY, 
 AdoEnums.CommandType.TEXT); 
 rstAuthors.getField("zip").getProperties(). 
 getItem("Optimize").setBoolean(true); // Create the index. 
 rstAuthors.find("zip = '94595'"); // Find Akiko Yokomoto. 
 System.out.println(rstAuthors.getField("au_fname").getString() + 
 " " + rstAuthors.getField("au_lname").getString() + " " + 
 rstAuthors.getField("address").getString() + " " + 
 rstAuthors.getField("city").getString() + " " + 
 rstAuthors.getField("state").getString() + " "); 
 rstAuthors.getField("zip").getProperties(). 
 getItem("Optimize").setBoolean(false); // Delete the index. 
 
 System.out.println("\nPress <Enter> to continue.."); 
 in.readLine(); 
 } 
 catch( AdoException ae ) 
 { 
 // Notify user of any errors that result from ADO. 
 
 // As passing a Recordset, check for null pointer first. 
 if (rstAuthors != null) 
 { 
 PrintProviderError(rstAuthors.getActiveConnection()); 
 } 
 else 
 { 
 System.out.println("Exception: " + ae.getMessage()); 
 } 
 } 
 
 // System read requires this catch. 
 catch( java.io.IOException je) 
 { 
 PrintIOError(je); 
 } 
 
 finally 
 { 
 // Cleanup objects before exit. 
 if (rstAuthors != null) 
 if (rstAuthors.getState() == 1) 
 rstAuthors.close(); 
 } 
 } 
 
 // PrintProviderError Function 
 
 static void PrintProviderError( Connection Cnn1 ) 
 { 
 // Print Provider errors from Connection object. 
 // ErrItem is an item object in the Connections Errors collection. 
 com.ms.wfc.data.Error ErrItem = null; 
 long nCount = 0; 
 int i = 0; 
 
 nCount = Cnn1.getErrors().getCount(); 
 
 // If there are any errors in the collection, print them. 
 if( nCount > 0); 
 { 
 // Collection ranges from 0 to nCount - 1 
 for (i = 0; i< nCount; i++) 
 { 
 ErrItem = Cnn1.getErrors().getItem(i); 
 System.out.println("\t Error number: " + ErrItem.getNumber() 
 + "\t" + ErrItem.getDescription() ); 
 } 
 } 
 
 } 
 
 // PrintIOError Function 
 
 static void PrintIOError( java.io.IOException je) 
 { 
 System.out.println("Error \n"); 
 System.out.println("\tSource = " + je.getClass() + "\n"); 
 System.out.println("\tDescription = " + je.getMessage() + "\n"); 
 } 
} 
// EndOptimizeJ 
```
