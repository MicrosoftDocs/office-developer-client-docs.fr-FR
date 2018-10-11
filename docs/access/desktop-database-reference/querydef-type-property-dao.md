---
title: QueryDef.Type Property (DAO)
TOCTitle: Type Property
ms:assetid: 03db891d-b958-7cf9-56c1-524d9ff2b9b5
ms:mtpsurl: https://msdn.microsoft.com/library/Ff844814(v=office.15)
ms:contentKeyID: 48542993
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 5129f4a50f056259c5b556c8e8ea408358718d3c
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470375"
---
# <a name="querydeftype-property-dao"></a><span data-ttu-id="fa7ef-102">QueryDef.Type Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="fa7ef-102">QueryDef.Type Property (DAO)</span></span>


<span data-ttu-id="fa7ef-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="fa7ef-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="fa7ef-p101">Définit ou renvoie une valeur qui indique le type opérationnel ou le type de données d'un objet. Type **Integer** en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="fa7ef-p101">Sets or returns a value that indicates the operational type or data type of an object. Read-only**Integer**.</span></span>

## <a name="syntax"></a><span data-ttu-id="fa7ef-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="fa7ef-106">Syntax</span></span>

<span data-ttu-id="fa7ef-107">*expression* . Type</span><span class="sxs-lookup"><span data-stu-id="fa7ef-107">*expression* .Type</span></span>

<span data-ttu-id="fa7ef-108">*expression* Variable qui représente un objet **QueryDef** .</span><span class="sxs-lookup"><span data-stu-id="fa7ef-108">*expression* A variable that represents a **QueryDef** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="fa7ef-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="fa7ef-109">Remarks</span></span>

<span data-ttu-id="fa7ef-110">Les paramètres et valeurs de retour possible d'un objet **QueryDef** sont décrits dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="fa7ef-110">For a **QueryDef** object, the possible settings and return values are shown in the following table.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="fa7ef-111">Constante</span><span class="sxs-lookup"><span data-stu-id="fa7ef-111">Constant</span></span></p></th>
<th><p><span data-ttu-id="fa7ef-112">Type de requête</span><span class="sxs-lookup"><span data-stu-id="fa7ef-112">Query type</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="fa7ef-113"><strong>dbQAction</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-113"><strong>dbQAction</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-114">Opération</span><span class="sxs-lookup"><span data-stu-id="fa7ef-114">Action</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="fa7ef-115"><strong>dbQAppend</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-115"><strong>dbQAppend</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-116">Ajout</span><span class="sxs-lookup"><span data-stu-id="fa7ef-116">Append</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="fa7ef-117"><strong>dbQCompound</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-117"><strong>dbQCompound</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-118">Composé</span><span class="sxs-lookup"><span data-stu-id="fa7ef-118">Compound</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="fa7ef-119"><strong>dbQCrosstab</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-119"><strong>dbQCrosstab</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-120">Analyse croisée</span><span class="sxs-lookup"><span data-stu-id="fa7ef-120">Crosstab</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="fa7ef-121"><strong>dbQDDL</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-121"><strong>dbQDDL</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-122">Définition de données</span><span class="sxs-lookup"><span data-stu-id="fa7ef-122">Data-definition</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="fa7ef-123"><strong>dbQDelete</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-123"><strong>dbQDelete</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-124">Delete</span><span class="sxs-lookup"><span data-stu-id="fa7ef-124">Delete</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="fa7ef-125"><strong>dbQMakeTable</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-125"><strong>dbQMakeTable</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-126">Création de table</span><span class="sxs-lookup"><span data-stu-id="fa7ef-126">Make-table</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="fa7ef-127"><strong>dbQProcedure</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-127"><strong>dbQProcedure</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-128">Procédure (espaces de travail ODBCDirect uniquement)</span><span class="sxs-lookup"><span data-stu-id="fa7ef-128">Procedure (ODBCDirect workspaces only)</span></span></p>

> [!NOTE]
> <P><span data-ttu-id="fa7ef-p102">[!REMARQUE] Les espaces de travail ODBCDirect ne sont pas pris en charge dans Microsoft Access 2013. Utilisez ADO si vous voulez accéder aux sources de données externes sans avoir recours au moteur de base de données Microsoft Access.</span><span class="sxs-lookup"><span data-stu-id="fa7ef-p102">ODBCDirect workspaces are not supported in Microsoft Access 2013. Use ADO if you want to access external data sources without using the Microsoft Access database engine.</span></span></P>


<p></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="fa7ef-131"><strong>dbQSelect</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-131"><strong>dbQSelect</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-132">Sélection</span><span class="sxs-lookup"><span data-stu-id="fa7ef-132">Select</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="fa7ef-133"><strong>dbQSetOperation</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-133"><strong>dbQSetOperation</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-134">Union</span><span class="sxs-lookup"><span data-stu-id="fa7ef-134">Union</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="fa7ef-135"><strong>dbQSPTBulk</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-135"><strong>dbQSPTBulk</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-136">Utilisé avec <strong>dbQSQLPassThrough</strong> pour spécifier une requête qui ne renvoie pas d’enregistrements (espaces de travail Microsoft Access uniquement).</span><span class="sxs-lookup"><span data-stu-id="fa7ef-136">Used with <strong>dbQSQLPassThrough</strong> to specify a query that doesn't return records (Microsoft Access workspaces only).</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="fa7ef-137"><strong>dbQSQLPassThrough</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-137"><strong>dbQSQLPassThrough</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-138">Requête SQL directe (espaces de travail Microsoft Access uniquement)</span><span class="sxs-lookup"><span data-stu-id="fa7ef-138">Pass-through (Microsoft Access workspaces only)</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="fa7ef-139"><strong>dbQUpdate</strong></span><span class="sxs-lookup"><span data-stu-id="fa7ef-139"><strong>dbQUpdate</strong></span></span></p></td>
<td><p><span data-ttu-id="fa7ef-140">Update</span><span class="sxs-lookup"><span data-stu-id="fa7ef-140">Update</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="fa7ef-141">Lorsque vous ajoutez un nouvel objet **[Field](field-object-dao.md)**, **[Parameter](parameter-object-dao.md)** ou **[Property](property-object-dao.md)** à la collection d'un objet **[Index](index-object-dao.md)**, **QueryDef**, **[Recordset](recordset-object-dao.md)** ou **[TableDef](tabledef-object-dao.md)**, une erreur se produit si la base de données sous-jacente ne prend pas en charge le type de données spécifié du nouvel objet.</span><span class="sxs-lookup"><span data-stu-id="fa7ef-141">When you append a new **[Field](field-object-dao.md)**, **[Parameter](parameter-object-dao.md)**, or **[Property](property-object-dao.md)** object to the collection of an **[Index](index-object-dao.md)**, **QueryDef**, **[Recordset](recordset-object-dao.md)**, or **[TableDef](tabledef-object-dao.md)** object, an error occurs if the underlying database doesn't support the data type specified for the new object.</span></span>
