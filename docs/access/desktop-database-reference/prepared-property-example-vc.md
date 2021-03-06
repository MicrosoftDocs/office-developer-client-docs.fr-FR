---
title: Prepared, propriété – Exemple (VC++)
TOCTitle: Prepared property example (VC++)
ms:assetid: 9b2d8037-e74d-5fbd-c56c-18187236b1b2
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249704(v=office.15)
ms:contentKeyID: 48546562
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: ac0009940b3c7917db82db38d604ef6d16712953
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32301427"
---
# <a name="prepared-property-example-vc"></a>Prepared, propriété – Exemple (VC++)


**S’applique à** : Access 2013, Office 2013

Cet exemple illustre la propriété [Prepared](prepared-property-ado.md) par l’ouverture de deux objets [Command](command-object-ado.md) : l’un préparé et l’autre non préparé.

```cpp 
 
// BeginPreparedCpp 
#import "c:\Program Files\Common Files\System\ADO\msado15.dll" no_namespace rename("EOF", "EndOfFile") 
 
#include <ole2.h> 
#include <stdio.h> 
#include <conio.h> 
#include <winbase.h> 
 
// Function declarations 
inline void TESTHR(HRESULT x) {if FAILED(x) _com_issue_error(x);}; 
void PreparedX(void); 
void PrintProviderError(_ConnectionPtr pConnection); 
void PrintComError(_com_error &e); 
 
/////////////////////////////////////////////////////////// 
// // 
// Main Function // 
// // 
/////////////////////////////////////////////////////////// 
 
void main() 
{ 
 if(FAILED(::CoInitialize(NULL))) 
 return; 
 
 PreparedX(); 
 
 ::CoUninitialize(); 
} 
 
/////////////////////////////////////////////////////////// 
// // 
// PreparedX Function // 
// // 
/////////////////////////////////////////////////////////// 
 
void PreparedX(void) 
{ 
 // Define ADO object pointers. 
 // Initialize pointers on define. 
 // These are in the ADODB:: namespace. 
 _ConnectionPtr pConnection =NULL; 
 _CommandPtr pCmd1 =NULL; 
 _CommandPtr pCmd2 =NULL; 
 
 // Define string variables. 
 _bstr_t strCnn("Provider='sqloledb';Data Source='MySqlServer';" 
 "Initial Catalog='pubs';Integrated Security='SSPI';"); 
 
 //Define Other Variables 
 HRESULT hr = S_OK; 
 
 try 
 { 
 // Open a connection. 
 TESTHR(pConnection.CreateInstance(__uuidof(Connection))); 
 pConnection->Open (strCnn, "", "", adConnectUnspecified); 
 
 _bstr_t strCmd ("SELECT title,type FROM titles ORDER BY type"); 
 
 // Create two command objects for the same 
 // command - one prepared and one not prepared. 
 TESTHR(pCmd1.CreateInstance(__uuidof(Command))); 
 pCmd1->ActiveConnection = pConnection; 
 pCmd1->CommandText = strCmd; 
 
 TESTHR(pCmd2.CreateInstance(__uuidof(Command))); 
 pCmd2->ActiveConnection = pConnection; 
 pCmd2->CommandText = strCmd; 
 pCmd2->PutPrepared(true); 
 
 // Set a timer,then execute the unprepared command 20 times. 
 DWORD sngStart=GetTickCount(); 
 
 for(int intLoop=1;intLoop<=20;intLoop++) 
 { 
 pCmd1->Execute(NULL,NULL,adCmdText); 
 } 
 DWORD sngEnd=GetTickCount(); 
 
 float sngNotPrepared = (float)(sngEnd - sngStart)/(float)1000; 
 
 // Reset the timer,then execute the prepared command 20 times 
 sngStart=GetTickCount(); 
 for(intLoop=1;intLoop<=20;intLoop++) 
 { 
 pCmd2->Execute(NULL,NULL,adCmdText); 
 } 
 sngEnd=GetTickCount(); 
 
 float sngPrepared = (float)(sngEnd - sngStart)/(float)1000; 
 
 // Display performance results 
 printf("\n\nPerformance Results:"); 
 printf("\n\nNot Prepared: %6.3f seconds",sngNotPrepared); 
 printf("\n\nPrepared: %6.3f seconds\n\n",sngPrepared); 
 } 
 catch (_com_error &e) 
 { 
 // Notify the user of errors if any. 
 // Pass a connection pointer accessed from the Connection. 
 PrintProviderError(pConnection); 
 PrintComError(e); 
 } 
 
 if (pConnection) 
 if (pConnection->State == adStateOpen) 
 pConnection->Close(); 
} 
 
/////////////////////////////////////////////////////////// 
// // 
// PrintProviderError Function // 
// // 
/////////////////////////////////////////////////////////// 
void PrintProviderError(_ConnectionPtr pConnection) 
{ 
 // Print Provider Errors from Connection object. 
 // pErr is a record object in the Connection's Error collection. 
 ErrorPtr pErr = NULL; 
 
 if( (pConnection->Errors->Count) > 0) 
 { 
 long nCount = pConnection->Errors->Count; 
 
 // Collection ranges from 0 to nCount -1. 
 for(long i = 0;i < nCount;i++) 
 { 
 pErr = pConnection->Errors->GetItem(i); 
 printf("\t Error number: %x\t%s", pErr->Number, 
 pErr->Description); 
 } 
 } 
} 
 
/////////////////////////////////////////////////////////// 
// // 
// PrintComError Function // 
// // 
/////////////////////////////////////////////////////////// 
void PrintComError(_com_error &e) 
{ 
 _bstr_t bstrSource(e.Source()); 
 _bstr_t bstrDescription(e.Description()); 
 
 // Print COM errors. 
 printf("Error\n"); 
 printf("\tCode = %08lx\n", e.Error()); 
 printf("\tCode meaning = %s\n", e.ErrorMessage()); 
 printf("\tSource = %s\n", (LPCSTR) bstrSource); 
 printf("\tDescription = %s\n", (LPCSTR) bstrDescription); 
} 
// EndPreparedCpp 
```

