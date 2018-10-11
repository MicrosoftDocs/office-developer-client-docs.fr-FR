---
title: BeginTrans, CommitTrans et RollbackTrans, méthodes-exemple (VC ++)
TOCTitle: BeginTrans, CommitTrans, and RollbackTrans Methods Example (VC++)
ms:assetid: b2e53b79-4f10-f9cc-59f7-61c9557e5ef2
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249853(v=office.15)
ms:contentKeyID: 48547182
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 3b9d003ff76295a88c26e29138af4ce2219db755
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471236"
---
# <a name="begintrans-committrans-and-rollbacktrans-methods-example-vc"></a><span data-ttu-id="2d8d0-102">BeginTrans, CommitTrans et RollbackTrans, méthodes - Exemple (VC++)</span><span class="sxs-lookup"><span data-stu-id="2d8d0-102">BeginTrans, CommitTrans, and RollbackTrans Methods Example (VC++)</span></span>


<span data-ttu-id="2d8d0-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="2d8d0-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="2d8d0-104">Cet exemple montre comment modifier le type de tous les livres de psychologie dans la table ***Titles*** de la base de données.</span><span class="sxs-lookup"><span data-stu-id="2d8d0-104">This example changes the book type of all psychology books in the ***Titles*** table of the database.</span></span> <span data-ttu-id="2d8d0-105">Une fois que la méthode [BeginTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) démarre une transaction qui isole toutes les modifications apportées à la table ***Titles*** , la méthode [CommitTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) enregistre les modifications.</span><span class="sxs-lookup"><span data-stu-id="2d8d0-105">After the [BeginTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) method starts a transaction that isolates all the changes made to the ***Titles*** table, the [CommitTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) method saves the changes.</span></span> <span data-ttu-id="2d8d0-106">Vous pouvez utiliser la méthode [RollbackTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) pour annuler les modifications que vous avez enregistré à l’aide de la méthode de [mise à jour](update-method-ado.md) .</span><span class="sxs-lookup"><span data-stu-id="2d8d0-106">You can use the [RollbackTrans](begintrans-committrans-and-rollbacktrans-methods-ado.md) method to undo changes that you saved using the [Update](update-method-ado.md) method.</span></span>

```cpp 
 
// BeginBeginTransCpp 
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" no_namespace rename("EOF", "EndOfFile") 
 
#include <stdio.h> 
#include <ole2.h> 
#include <conio.h> 
#include <assert.h> 
#include <malloc.h> 
#include "BeginTransX.h" 
 
inline void TESTHR(HRESULT x) {if FAILED(x) _com_issue_error(x);}; 
void BeginTransX(void); 
void PrintProviderError(_ConnectionPtr pConnection); 
 
void main() 
{ 
 if(FAILED(::CoInitialize(NULL))) 
 return; 
 
 BeginTransX(); 
 
 //Wait here for user to see the output. 
 printf("\nPress any key to continue..."); 
 getch(); 
 
 ::CoUninitialize(); 
} 
 
/////////////////////////////////////////////////////////// 
// // 
// BeginTransX Function // 
// // 
/////////////////////////////////////////////////////////// 
 
void BeginTransX() 
{ 
 // Define ADO object pointers. 
 // Initialize pointers on define. 
 // These are in the ADODB:: namespace. 
 _RecordsetPtr rstTitles = NULL; 
 _ConnectionPtr pConnection = NULL; 
 
 //Define Other Variables 
 HRESULT hr = S_OK; 
 IADORecordBinding *picRs = NULL; //Interface Pointer declared... 
 CTitlesRs titlrs; 
 _bstr_t strTitle; 
 _bstr_t strMessage; 
 LPSTR p_TempStr = NULL; 
 char chKey; 
 int i = 0; 
 
 try 
 { 
 // open connection. 
 _bstr_t strCnn("Provider='sqloledb';Data Source='MySqlServer';" 
 "Initial Catalog='pubs';Integrated Security='SSPI';"); 
 
 TESTHR(pConnection.CreateInstance(__uuidof(Connection))); 
 
 TESTHR(pConnection->Open(strCnn,"","",adConnectUnspecified)); 
 
 rstTitles.CreateInstance(__uuidof(Recordset)); 
 
 rstTitles->CursorType = adOpenDynamic; 
 rstTitles->LockType = adLockPessimistic; 
 
 // open Titles table 
 TESTHR(rstTitles->Open("titles", 
 _variant_t((IDispatch*)pConnection,true), 
 adOpenDynamic, adLockPessimistic,adCmdTable)); 
 
 rstTitles->MoveFirst(); 
 pConnection->BeginTrans(); 
 
 //Open an IADORecordBinding interface pointer which 
 //we'll use for Binding Recordset to a class 
 TESTHR(rstTitles->QueryInterface( 
 __uuidof(IADORecordBinding), (LPVOID*)&picRs)); 
 
 //Bind the Recordset to a C++ Class here 
 TESTHR(picRs->BindToRecordset(&titlrs)); 
 
 // Loop through recordset and ask user if he wants 
 // to change the type for a specified title. 
 // Allocate memory to p_TempStr string pointer. 
 p_TempStr = (LPSTR) malloc(sizeof(titlrs.m_szT_type)); 
 
 // Check for null string. 
 assert(p_TempStr != NULL); 
 while (VARIANT_FALSE == rstTitles->EndOfFile) 
 { 
 // Set the final character of the destination string to NULL. 
 p_TempStr[sizeof(titlrs.m_szT_type)-1] = '\0'; 
 // The source string will get truncated if its length is 
 // longer than the length of the destination string minus one. 
 strncpy(p_TempStr,strtok(titlrs.m_szT_type," "),sizeof(titlrs.m_szT_type)-1); 
 
 // Compare type with psychology 
 if (!strcmp(p_TempStr,"psychology")) 
 { 
 strTitle = titlrs.m_szT_title; 
 strMessage = "Title: " + strTitle + 
 "\n Change type to Self help?(y/n)"; 
 
 // Change the title for specified employee 
 printf("%s\n",(LPCSTR)strMessage); 
 do 
 { 
 chKey = getch(); 
 }while(chKey != 'y' && chKey !='n'); 
 if(chKey == 'y') 
 { 
 // Set the final character of the destination string to NULL. 
 titlrs.m_szT_type[sizeof(titlrs.m_szT_type)-1] = '\0'; 
 // Copy "self_help" title field. 
 // The source string will get truncated if its length is 
 // longer than the length of the destination string minus one. 
 strncpy(titlrs.m_szT_type,"self_help", sizeof(titlrs.m_szT_type)-1); 
 picRs->Update(&titlrs); 
 } 
 } 
 rstTitles->MoveNext(); 
 } 
 // Ask if the User wants to commit to all the 
 // changes made above 
 printf("\n\n Save all changes(y/n)?"); 
 do 
 { 
 chKey = getch(); 
 }while(chKey != 'y' && chKey !='n'); 
 
 if(chKey == 'y') 
 
 // Save the changes to the title table 
 pConnection->CommitTrans(); 
 else 
 // Unsave the changes to the title table 
 pConnection->RollbackTrans(); 
 
 // Print current data in recordset. 
 rstTitles->Requery(0); 
 
 // Move to the first record of the title table 
 rstTitles->MoveFirst(); 
 printf("\n\nPress any key to continue..."); 
 getch(); 
 
 // Clear the screen for the next display 
 //system("cls"); 
 
 // Open IADORecordBinding interface pointer again 
 // for binding Recordset to a class. 
 TESTHR(rstTitles->QueryInterface( 
 __uuidof(IADORecordBinding), 
 (LPVOID*)&picRs)); 
 
 // Rebind the Recordset to a C++ Class. 
 TESTHR(picRs->BindToRecordset(&titlrs)); 
 
 while (!rstTitles->EndOfFile) 
 { 
 i= i+1; 
 if (i % 23 == 0) 
 { 
 printf("\nPress any key to continue..."); 
 getch(); 
 
 //Clear the screen for the next display 
 //system("cls"); 
 } 
 printf("%s - %s\n",titlrs.m_szT_title,titlrs.m_szT_type); 
 rstTitles->MoveNext(); 
 } 
 // Restore original data becasue this is a demonstration. 
 rstTitles->MoveFirst(); 
 
 while (VARIANT_FALSE == rstTitles->EndOfFile) 
 { 
 // Set the final character of the destination string to NULL. 
 p_TempStr[sizeof(titlrs.m_szT_type)-1] = '\0'; 
 // The source string will get truncated if its length is 
 // longer than the length of the destination string minus one. 
 strncpy(p_TempStr,titlrs.m_szT_type,sizeof(titlrs.m_szT_type)-1); 
 p_TempStr = strtok(p_TempStr," "); 
 
 if (!strcmp(p_TempStr,"self_help")) 
 { 
 // Set the final character of the destination string to NULL. 
 titlrs.m_szT_type[sizeof(titlrs.m_szT_type)-1] = '\0'; 
 // The source string will get truncated if its length is 
 // longer than the length of the destination string minus one. 
 strncpy(titlrs.m_szT_type,"psychology",sizeof(titlrs.m_szT_type)-1); 
 picRs->Update(&titlrs); 
 } 
 rstTitles->MoveNext(); 
 } 
 } 
 catch(_com_error &e) 
 { 
 // Notify the user of errors if any. 
 _bstr_t bstrSource(e.Source()); 
 _bstr_t bstrDescription(e.Description()); 
 
 PrintProviderError(pConnection); 
 
 printf("Source : %s\n",(LPCSTR)bstrSource); 
 printf("Description : %s\n",(LPCSTR)bstrDescription); 
 } 
 
 // Deallocate the memory 
 if (p_TempStr) 
 free(p_TempStr); 
 // Clean up objects before exit. 
 //Release the IADORecordset Interface here 
 if (picRs) 
 picRs->Release(); 
 
 if (rstTitles) 
 if (rstTitles->State == adStateOpen) 
 rstTitles->Close(); 
 if (pConnection) 
 if (pConnection->State == adStateOpen) 
 pConnection->Close(); 
}; 
 
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
 long nCount = 0; 
 long i = 0; 
 
 if( (pConnection->Errors->Count) > 0) 
 { 
 nCount = pConnection->Errors->Count; 
 // Collection ranges from 0 to nCount -1. 
 for(i = 0; i < nCount; i++) 
 { 
 pErr = pConnection->Errors->GetItem(i); 
 printf("Error number: %x\t%s\n", pErr->Number,(LPCSTR) pErr->Description); 
 } 
 } 
} 
// EndBeginTransCpp 
```

<span data-ttu-id="2d8d0-107">**BeginTransX.h**</span><span class="sxs-lookup"><span data-stu-id="2d8d0-107">**BeginTransX.h**</span></span>

```cpp 
 
// BeginBeginTransH 
#include "icrsint.h" 
 
//This Class extracts only title and type 
class CTitlesRs : public CADORecordBinding 
{ 
BEGIN_ADO_BINDING(CTitlesRs) 
 //Column title is the 2nd field in the recordset 
 ADO_VARIABLE_LENGTH_ENTRY2(2, adVarChar, m_szT_title, 
 sizeof(m_szT_title), lT_titleStatus, FALSE) 
 
 //Column type is the 3rd field in the recordset 
 ADO_VARIABLE_LENGTH_ENTRY2(3, adVarChar, m_szT_type, 
 sizeof(m_szT_type), lT_typeStatus, TRUE) 
END_ADO_BINDING() 
 
public: 
 CHAR m_szT_title[150]; 
 ULONG lT_titleStatus; 
 CHAR m_szT_type[40]; 
 ULONG lT_typeStatus; 
}; 
// EndBeginTransH 
```
