---
title: ADORecordsetConstruction, interface (ADO)
TOCTitle: ADORecordsetConstruction Interface (ADO)
ms:assetid: 2b53aa6e-3b6f-a996-3967-534215fd586c
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249060(v=office.15)
ms:contentKeyID: 48543926
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 0b99fcfe4fadc3dddf5937ba1043875de43e88d6
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471398"
---
# <a name="adorecordsetconstruction-interface-ado"></a><span data-ttu-id="e91da-102">ADORecordsetConstruction, interface (ADO)</span><span class="sxs-lookup"><span data-stu-id="e91da-102">ADORecordsetConstruction Interface (ADO)</span></span>


<span data-ttu-id="e91da-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="e91da-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="e91da-104">L'interface **ADORecordsetConstruction** permet de créer un objet **Recordset** ADO à partir d'un objet **Rowset** OLE DB dans une application C/C++.</span><span class="sxs-lookup"><span data-stu-id="e91da-104">The **ADORecordsetConstruction** interface is used to construct an ADO **Recordset** object from an OLE DB **Rowset** object in a C/C++ application.</span></span>

<span data-ttu-id="e91da-105">Cette interface prend en charge les propriétés ci-après :</span><span class="sxs-lookup"><span data-stu-id="e91da-105">This interface supports the following properties:</span></span>

## <a name="properties"></a><span data-ttu-id="e91da-106">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e91da-106">Properties</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="e91da-107"><a href="chapter-property-ado.md">Chapitre</a></span><span class="sxs-lookup"><span data-stu-id="e91da-107"><a href="chapter-property-ado.md">Chapter</a></span></span></p></td>
<td><p><span data-ttu-id="e91da-108">En lecture-écriture.</span><span class="sxs-lookup"><span data-stu-id="e91da-108">Read/Write.</span></span><br />
<span data-ttu-id="e91da-109">Obtient ou définit un objet OLE DB <strong>chapitre</strong> à partir de/sur cet objet ADO <strong>Recordset</strong> .</span><span class="sxs-lookup"><span data-stu-id="e91da-109">Gets/sets an OLE DB <strong>Chapter</strong> object from/on this ADO <strong>Recordset</strong> object.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="e91da-110"><a href="rowposition-property-ado.md">RowPosition</a></span><span class="sxs-lookup"><span data-stu-id="e91da-110"><a href="rowposition-property-ado.md">RowPosition</a></span></span></p></td>
<td><p><span data-ttu-id="e91da-111">En lecture-écriture.</span><span class="sxs-lookup"><span data-stu-id="e91da-111">Read/Write.</span></span><br />
<span data-ttu-id="e91da-112">Obtient ou définit un objet OLE DB <strong>RowPosition</strong> à partir de/sur cet objet ADO <strong>Recordset</strong> .</span><span class="sxs-lookup"><span data-stu-id="e91da-112">Gets/sets an OLE DB <strong>RowPosition</strong> object from/on this ADO <strong>Recordset</strong> object.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="e91da-113"><a href="rowset-property-ado.md">Ensemble de lignes</a></span><span class="sxs-lookup"><span data-stu-id="e91da-113"><a href="rowset-property-ado.md">Rowset</a></span></span></p></td>
<td><p><span data-ttu-id="e91da-114">En lecture-écriture.</span><span class="sxs-lookup"><span data-stu-id="e91da-114">Read/Write.</span></span><br />
<span data-ttu-id="e91da-115">Obtient ou définit un objet OLE DB <strong>Rowset</strong> à partir de/sur cet objet ADO <strong>Recordset</strong> .</span><span class="sxs-lookup"><span data-stu-id="e91da-115">Gets/sets an OLE DB <strong>Rowset</strong> object from/on this ADO <strong>Recordset</strong> object.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="methods"></a><span data-ttu-id="e91da-116">Méthodes</span><span class="sxs-lookup"><span data-stu-id="e91da-116">Methods</span></span>

<span data-ttu-id="e91da-117">Aucune.</span><span class="sxs-lookup"><span data-stu-id="e91da-117">None.</span></span>

## <a name="events"></a><span data-ttu-id="e91da-118">Événements</span><span class="sxs-lookup"><span data-stu-id="e91da-118">Events</span></span>

<span data-ttu-id="e91da-119">Aucune.</span><span class="sxs-lookup"><span data-stu-id="e91da-119">None.</span></span>

## <a name="remarks"></a><span data-ttu-id="e91da-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="e91da-120">Remarks</span></span>

<span data-ttu-id="e91da-121">Étant donné un objet OLE DB **Rowset** (pRowset), la construction d’un objet de ADO **Recordset** (), la construction d’un objet **Recordset** ADO les chiffres d’objet (adoRs) pour les trois opérations ci-après :</span><span class="sxs-lookup"><span data-stu-id="e91da-121">Given an OLE DB **Rowset** object (pRowset ), the construction of an ADO **Recordset** object (), the construction of an ADO **Recordset** object (adoRs ) amounts to the following three basic operations:</span></span>

1. <span data-ttu-id="e91da-122">Créez un objet **Recordset** ADO :</span><span class="sxs-lookup"><span data-stu-id="e91da-122">Create an ADO **Recordset** object:</span></span>
    
   ```vb
    Recordset20Ptr adoRs;
    adoRs.CreateInstance(__uuidof(Recordset));
   ```
2. <span data-ttu-id="e91da-123">Interrogez l'interface **IADORecordsetConstruction** sur l'objet **Recordset**:</span><span class="sxs-lookup"><span data-stu-id="e91da-123">Query the **IADORecordsetConstruction** interface on the **Recordset** object:</span></span>

   ```vb    
    adoRecordsetConstructionPtr adoRsConstruct=NULL;
    adoRs->QueryInterface(__uuidof(ADORecordsetConstruction),
         (void**)&adoRsConstruct);
   ```

3. <span data-ttu-id="e91da-124">Appelez le IADORecordsetConstruction::put\_méthode de la propriété Rowset pour définir l’objet Rowset OLE DB sur l’objet Recordset ADO :</span><span class="sxs-lookup"><span data-stu-id="e91da-124">Call the IADORecordsetConstruction::put\_Rowset property method to set the OLE DB Rowset object on the ADO Recordset object:</span></span>

   ```vb     
    IUnknown *pUnk=NULL;
    pRowset->QueryInterface(IID_IUnknown, (void**)&pUnk);
    adoRsConstruct->put_Rowset(pUnk);
   ```
<span data-ttu-id="e91da-125">L’objet qui en résulte représente maintenant l’objet de l’objet **Recordset** ADO construit à partir de l’objet OLE DB **Rowset** .</span><span class="sxs-lookup"><span data-stu-id="e91da-125">The resultant object now represents the ADO **Recordset** object constructed from the OLE DB **Rowset** object.</span></span>

<span data-ttu-id="e91da-126">Vous pouvez également créer un objet **Recordset** ADO à partir d'un objet **Chapter** ou **RowPosition** OLE DB.</span><span class="sxs-lookup"><span data-stu-id="e91da-126">You can also construct an ADO **Recordset** object from an OLE DB **Chapter** or **RowPosition** object.</span></span>

## <a name="requirements"></a><span data-ttu-id="e91da-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e91da-127">Requirements</span></span>

- <span data-ttu-id="e91da-128">**Version :** ADO 2.0 et supérieure</span><span class="sxs-lookup"><span data-stu-id="e91da-128">**Version:** ADO 2.0 and later</span></span>

- <span data-ttu-id="e91da-129">**Bibliothèque :** msado15.dll</span><span class="sxs-lookup"><span data-stu-id="e91da-129">**Library:** msado15.dll</span></span>

- <span data-ttu-id="e91da-130">**UUID :** 00000283-0000-0010-8000-00AA006D2EA4</span><span class="sxs-lookup"><span data-stu-id="e91da-130">**UUID:** 00000283-0000-0010-8000-00AA006D2EA4</span></span>
