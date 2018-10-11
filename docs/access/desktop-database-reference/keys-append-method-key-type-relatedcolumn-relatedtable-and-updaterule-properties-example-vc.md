---
title: Keys Append, méthode, Key Type, RelatedColumn propriétés-exemple (VC ++)
TOCTitle: Keys Append Method, Key Type, RelatedColumn, RelatedTable and UpdateRule Properties Example (VC++)
ms:assetid: d0784eb5-94aa-ef62-c26f-3d0980485990
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250041(v=office.15)
ms:contentKeyID: 48547840
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 0181f93beee3546690ea2a4a806563b2cc220bc7
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469742"
---
# <a name="keys-append-method-key-type-relatedcolumn-relatedtable-and-updaterule-properties-example-vc"></a><span data-ttu-id="781fb-102">Keys Append, méthode - Exemple de propriétés Key Type, RelatedColumn, RelatedTable et UpdateRule (VC++)</span><span class="sxs-lookup"><span data-stu-id="781fb-102">Keys Append Method, Key Type, RelatedColumn, RelatedTable and UpdateRule Properties Example (VC++)</span></span>


<span data-ttu-id="781fb-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="781fb-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="781fb-p101">Le code suivant montre comment créer une clé étrangère. Il suppose l'existence de deux tables (Customers et Orders).</span><span class="sxs-lookup"><span data-stu-id="781fb-p101">The following code demonstrates how to create a new foreign key. It assumes two tables (Customers and Orders) exist.</span></span>

```cpp 
 
// BeginCreateKeyCpp 
#import "c:\Program Files\Common Files\system\ado\msadox.dll" no_namespace 
#import "c:\Program Files\Common Files\system\ado\msado15.dll" 
 
#include "iostream.h" 
#include "stdio.h" 
#include "conio.h" 
 
//Function declarations 
inline void TESTHR(HRESULT x) {if FAILED(x) _com_issue_error(x);}; 
void CreateKeyX(void); 
 
////////////////////////////////////////////////////////// 
// // 
// Main Function // 
// // 
////////////////////////////////////////////////////////// 
void main() 
{ 
 if(FAILED(::CoInitialize(NULL))) 
 return; 
 
 CreateKeyX(); 
 
 ::CoUninitialize(); 
} 
 
////////////////////////////////////////////////////////// 
// // 
// CreateKeyX Function // 
// // 
////////////////////////////////////////////////////////// 
void CreateKeyX(void) 
{ 
 HRESULT hr = S_OK; 
 
 // Define ADOX object pointers. 
 // Initialize pointers on define. 
 // These are in the ADOX:: namespace. 
 _KeyPtr m_pKeyForeign = NULL; 
 _CatalogPtr m_pCatalog = NULL; 
 
 //Define other variables 
 _bstr_t strcnn("Provider='Microsoft.JET.OLEDB.4.0';" 
 "Data source = 'c:\\Program Files\\Microsoft Office\\" 
 "Office\\Samples\\Northwind.mdb';"); 
 try 
 { 
 TESTHR(hr = m_pKeyForeign.CreateInstance(__uuidof(Key))); 
 TESTHR(hr = m_pCatalog.CreateInstance(__uuidof (Catalog))); 
 m_pCatalog->PutActiveConnection(strcnn); 
 
 // Define the foreign key. 
 m_pKeyForeign->Name = "CustOrder"; 
 m_pKeyForeign->Type = adKeyForeign; 
 m_pKeyForeign->RelatedTable = "Customers"; 
 m_pKeyForeign->Columns->Append("CustomerId",adVarWChar,0); 
 m_pKeyForeign->Columns->GetItem("CustomerId")->RelatedColumn = 
 "CustomerId"; 
 m_pKeyForeign->UpdateRule = adRICascade; 
 
 // To pass as column parameter to Key's Apppend method 
 _variant_t vOptional; 
 vOptional.vt = VT_ERROR; 
 vOptional.scode = DISP_E_PARAMNOTFOUND; 
 
 // Append the foreign key. 
 m_pCatalog->Tables->GetItem("Orders")->Keys-> 
 Append(_variant_t((IDispatch *)m_pKeyForeign), 
 adKeyPrimary,vOptional,L"",L""); 
 printf("Foreign key 'CustOrder' is added.\n"); 
 
 // Delete the key as this is a demonstration. 
 m_pCatalog->Tables->GetItem("Orders")->Keys-> 
 Delete(m_pKeyForeign->Name); 
 printf("Foreign key 'CustOrder' is deleted.\n"); 
 } 
 
 catch(_com_error &e) 
 { 
 // Notify the user of errors if any. 
 _bstr_t bstrSource(e.Source()); 
 _bstr_t bstrDescription(e.Description()); 
 
 printf("\n\tSource : %s \n\tdescription : %s \n ", 
 (LPCSTR)bstrSource,(LPCSTR)bstrDescription); 
 } 
 
 catch(...) 
 { 
 cout << "Error occured in include files...."<< endl; 
 } 
} 
// EndCreateKeyCpp 
```
