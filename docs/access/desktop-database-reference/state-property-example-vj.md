---
title: State, propriété - Exemple (VJ++)
TOCTitle: State Property Example (VJ++)
ms:assetid: 7de6b4c1-b761-4060-7d97-6207542c202d
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249529(v=office.15)
ms:contentKeyID: 48545869
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 125a8caa3afc210c08d8ea1a4a8cea27e22a67a8
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469476"
---
# <a name="state-property-example-vj"></a><span data-ttu-id="6bee1-102">State, propriété - Exemple (VJ++)</span><span class="sxs-lookup"><span data-stu-id="6bee1-102">State Property Example (VJ++)</span></span>


<span data-ttu-id="6bee1-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="6bee1-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="6bee1-104">Cet exemple utilise la propriété [State](state-property-ado.md) pour afficher un message pendant que des connexions asynchrones s'ouvrent et que des commandes asynchrones s'exécutent.</span><span class="sxs-lookup"><span data-stu-id="6bee1-104">This example uses the [State](state-property-ado.md) property to display a message while asynchronous connections are opening and asynchronous commands are executing.</span></span>

```java 
 
// BeginStateJ 
import com.ms.wfc.data.*; 
import java.io.*; 
 
public class StateX 
{ 
 // The main entry point of the application. 
 public static void main (String[] args) 
 { 
 StateX(); 
 System.exit(0); 
 } 
 // StateX Function 
 static void StateX() 
 { 
 // Define ADO Objects. 
 Connection cnn1 = null; 
 Connection cnn2 = null; 
 Command cmdChange = null; 
 Command cmdRestore = null; 
 
 // Declarations. 
 String strCnn = "Provider='sqloledb';Data Source='MySqlServer';"+ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';"; 
 BufferedReader in = 
 new BufferedReader(new InputStreamReader(System.in)); 
 
 try 
 { 
 // Open two Asynchronous connections, displaying 
 // a message while connecting. 
 cnn1 = new Connection(); 
 cnn2 = new Connection(); 
 
 cnn1.open(strCnn,"","",AdoEnums.ConnectOption.ASYNCCONNECT); 
 while(cnn1.getState()==AdoEnums.ObjectState.CONNECTING) 
 System.out.println("Opening the first connection...."); 
 
 cnn2.open(strCnn,"","",AdoEnums.ConnectOption.ASYNCCONNECT); 
 while(cnn2.getState()==AdoEnums.ObjectState.CONNECTING) 
 System.out.println("Opening the second connection...."); 
 
 // Create two command Objects. 
 cmdChange = new Command(); 
 cmdChange.setActiveConnection(cnn1); 
 cmdChange.setCommandText("UPDATE Titles SET type = 'self_help'"+ 
 "WHERE type = 'psychology'"); 
 
 cmdRestore = new Command(); 
 cmdRestore.setActiveConnection(cnn2); 
 cmdRestore.setCommandText( 
 "UPDATE Titles SET type = 'psychology'"+ 
 "WHERE type = 'self_help'"); 
 
 // Executing the commands, displaying a message 
 // while they are executing. 
 cmdChange.execute(null,AdoEnums.ExecuteOption.ASYNCEXECUTE); 
 
 while(cmdChange.getState() == AdoEnums.ObjectState.EXECUTING) 
 System.out.println("Change command executing...."); 
 
 cmdRestore.execute(null,AdoEnums.ExecuteOption.ASYNCEXECUTE); 
 
 while(cmdRestore.getState() == AdoEnums.ObjectState.EXECUTING) 
 System.out.println("Restore command executing...."); 
 
 System.out.println("Press <Enter> to continue.."); 
 in.readLine(); 
 } 
 // System read requires this catch. 
 catch(java.io.IOException je) 
 { 
 PrintIOError(je); 
 } 
 catch(AdoException ae) 
 { 
 // Notify user of any errors resulting from ADO. 
 
 // As passing a connection, check for null pointer first. 
 if(cnn1 != null) 
 { 
 System.out.println("The error has occured in cnn1:"); 
 PrintProviderError(cnn1); 
 } 
 else if(cnn2 != null) 
 { 
 System.out.println("The error has occured in cnn2:"); 
 PrintProviderError(cnn2); 
 } 
 else 
 { 
 System.out.println("Exception: "+ ae.getMessage()); 
 } 
 } 
 
 finally 
 { 
 // Cleanup objects before exit. 
 if (cnn1 != null) 
 if (cnn1.getState() == 1) 
 cnn1.close(); 
 if (cnn2 != null) 
 if (cnn2.getState() == 1) 
 cnn2.close(); 
 } 
 
 } 
 // PrintProviderError Function 
 static void PrintProviderError(Connection cnn1) 
 { 
 // Print Provider Errors from Connection Object. 
 // ErrItem is an item object in the Connections Errors Collection. 
 com.ms.wfc.data.Error ErrItem = null; 
 long nCount = 0; 
 int i = 0; 
 
 nCount = cnn1.getErrors().getCount(); 
 
 // If there are any errors in the collection, print them. 
 if ( nCount > 0) 
 { 
 // Collection ranges from 0 to nCount-1. 
 for ( i=0;i<nCount; i++) 
 { 
 ErrItem = cnn1.getErrors().getItem(i); 
 System.out.println("\t Error Number: " + ErrItem.getNumber() 
 + "\t" + ErrItem.getDescription()); 
 } 
 } 
 } 
 // PrintIOError Function 
 static void PrintIOError(java.io.IOException je) 
 { 
 System.out.println("Error: \n"); 
 System.out.println("\t Source: " + je.getClass() + "\n"); 
 System.out.println("\t Description: "+ je.getMessage() + "\n"); 
 } 
} 
// EndStateJ 
```
