---
title: Field.GetChunk Method (DAO)
TOCTitle: GetChunk Method
ms:assetid: b8984e79-54f7-8052-85a3-d12033daf7a1
ms:mtpsurl: https://msdn.microsoft.com/library/Ff822448(v=office.15)
ms:contentKeyID: 48547328
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052871
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 2931dcda3c344163f7286222d354f9303966ba45
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472061"
---
# <a name="fieldgetchunk-method-dao"></a><span data-ttu-id="3c76c-102">Field.GetChunk Method (DAO)</span><span class="sxs-lookup"><span data-stu-id="3c76c-102">Field.GetChunk Method (DAO)</span></span>


<span data-ttu-id="3c76c-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="3c76c-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="3c76c-104">Renvoie tout ou partie du contenu d'un objet \*\*\*\*Field\*\*\*\* de type [Memo](field-object-dao.md) ou **Long Binary** appartenant à la collection **[Fields](fields-collection-dao.md)** d'un objet **[Recordset](recordset-object-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="3c76c-104">Returns all or a portion of the contents of a **Memo** or **Long Binary** **[Field](field-object-dao.md)** object in the **[Fields](fields-collection-dao.md)** collection of a **[Recordset](recordset-object-dao.md)** object.</span></span>

## <a name="syntax"></a><span data-ttu-id="3c76c-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="3c76c-105">Syntax</span></span>

<span data-ttu-id="3c76c-106">*expression* . GetChunk (***décalage***, ***octets***)</span><span class="sxs-lookup"><span data-stu-id="3c76c-106">*expression* .GetChunk(***Offset***, ***Bytes***)</span></span>

<span data-ttu-id="3c76c-107">*expression* Variable qui représente un objet **Field** .</span><span class="sxs-lookup"><span data-stu-id="3c76c-107">*expression* A variable that represents a **Field** object.</span></span>

### <a name="parameters"></a><span data-ttu-id="3c76c-108">Paramètres</span><span class="sxs-lookup"><span data-stu-id="3c76c-108">Parameters</span></span>

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="3c76c-109">Name</span><span class="sxs-lookup"><span data-stu-id="3c76c-109">Name</span></span></p></th>
<th><p><span data-ttu-id="3c76c-110">Obligatoire/Facultatif</span><span class="sxs-lookup"><span data-stu-id="3c76c-110">Required/Optional</span></span></p></th>
<th><p><span data-ttu-id="3c76c-111">Type de données</span><span class="sxs-lookup"><span data-stu-id="3c76c-111">Data Type</span></span></p></th>
<th><p><span data-ttu-id="3c76c-112">Description</span><span class="sxs-lookup"><span data-stu-id="3c76c-112">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="3c76c-113">Offset</span><span class="sxs-lookup"><span data-stu-id="3c76c-113">Offset</span></span></p></td>
<td><p><span data-ttu-id="3c76c-114">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="3c76c-114">Required</span></span></p></td>
<td><p><span data-ttu-id="3c76c-115"><strong>Entier long</strong></span><span class="sxs-lookup"><span data-stu-id="3c76c-115"><strong>Long</strong></span></span></p></td>
<td><p><span data-ttu-id="3c76c-116">Nombre d'octets à ignorer avant que la copie ne commence.</span><span class="sxs-lookup"><span data-stu-id="3c76c-116">The number of bytes to skip before copying begins.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3c76c-117">Octets</span><span class="sxs-lookup"><span data-stu-id="3c76c-117">Bytes</span></span></p></td>
<td><p><span data-ttu-id="3c76c-118">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="3c76c-118">Required</span></span></p></td>
<td><p><span data-ttu-id="3c76c-119"><strong>Entier long</strong></span><span class="sxs-lookup"><span data-stu-id="3c76c-119"><strong>Long</strong></span></span></p></td>
<td><p><span data-ttu-id="3c76c-120">Nombre d'octets que vous souhaitez renvoyer.</span><span class="sxs-lookup"><span data-stu-id="3c76c-120">The number of bytes you want to return.</span></span></p></td>
</tr>
</tbody>
</table>


### <a name="return-value"></a><span data-ttu-id="3c76c-121">Valeur renvoyée</span><span class="sxs-lookup"><span data-stu-id="3c76c-121">Return Value</span></span>

<span data-ttu-id="3c76c-122">Variante</span><span class="sxs-lookup"><span data-stu-id="3c76c-122">Variant</span></span>

## <a name="remarks"></a><span data-ttu-id="3c76c-123">Remarques</span><span class="sxs-lookup"><span data-stu-id="3c76c-123">Remarks</span></span>

<span data-ttu-id="3c76c-p101">Les octets renvoyés par **GetChunk** sont affectés à une variable. Utilisez **GetChunk** pour renvoyer une partie de la valeur totale des données à la fois. Vous pouvez avoir recours à **[AppendChunk](field-appendchunk-method-dao.md)** pour reconstituer les différentes parties.</span><span class="sxs-lookup"><span data-stu-id="3c76c-p101">The bytes returned by **GetChunk** are assigned to variable. Use **GetChunk** to return a portion of the total data value at a time. You can use the **[AppendChunk](field-appendchunk-method-dao.md)** method to reassemble the pieces.</span></span>

<span data-ttu-id="3c76c-127">Si le décalage est égal à 0, **GetChunk** commence la copie à partir du premier octet du champ.</span><span class="sxs-lookup"><span data-stu-id="3c76c-127">If offset is 0, **GetChunk** begins copying from the first byte of the field.</span></span>

<span data-ttu-id="3c76c-128">Si la valeur du paramètre NbOctets est supérieur au nombre d’octets dans le champ, **GetChunk** renvoie le nombre réel d’octets restant dans le champ.</span><span class="sxs-lookup"><span data-stu-id="3c76c-128">If numbytes is greater than the number of bytes in the field, **GetChunk** returns the actual number of remaining bytes in the field.</span></span>


> [!NOTE]
> <P><span data-ttu-id="3c76c-p102">[!REMARQUE] Utilisez un champ de type <STRONG>Memo</STRONG> pour du texte et placez les données binaires uniquement dans des champs de type <STRONG>Long Binary</STRONG>. Sinon, vous n'obtiendrez pas les résultats escomptés.</span><span class="sxs-lookup"><span data-stu-id="3c76c-p102">Use a <STRONG>Memo</STRONG> field for text, and put binary data only in <STRONG>Long Binary</STRONG> fields. Doing otherwise will cause undesirable results.</span></span></P>



## <a name="example"></a><span data-ttu-id="3c76c-131">Exemple</span><span class="sxs-lookup"><span data-stu-id="3c76c-131">Example</span></span>

<span data-ttu-id="3c76c-p103">Cet exemple utilise les méthodes **AppendChunk** et **GetChunk** pour remplir un champ d'objet OLE avec des données issues d'un autre enregistrement, 32 Ko à la fois. Dans une application réelle, un utilisateur peut avoir recours à une procédure de ce type pour copier un enregistrement d'employé (y compris sa photo) d'une table à une autre. Dans le cadre de cet exemple, l'enregistrement est simplement recopié dans la même table. Notez que toute la manipulation des segments a lieu dans une seule séquence AddNew-Update.</span><span class="sxs-lookup"><span data-stu-id="3c76c-p103">This example uses the **AppendChunk** and **GetChunk** methods to fill an OLE object field with data from another record, 32K at a time. In a real application, one might use a procedure like this to copy an employee record (including the employee's photo) from one table to another. In this example, the record is simply being copied back to same table. Note that all the chunk manipulation takes place within a single AddNew-Update sequence.</span></span>

```vb
    Sub AppendChunkX() 
     
     Dim dbsNorthwind As Database 
     Dim rstEmployees As Recordset 
     Dim rstEmployees2 As Recordset 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     
     ' Open two recordsets from the Employees table. 
     Set rstEmployees = _ 
     dbsNorthwind.OpenRecordset("Employees", _ 
     dbOpenDynaset) 
     Set rstEmployees2 = rstEmployees.Clone 
     
     ' Add a new record to the first Recordset and copy the 
     ' data from a record in the second Recordset. 
     With rstEmployees 
     .AddNew 
     !FirstName = rstEmployees2!FirstName 
     !LastName = rstEmployees2!LastName 
     CopyLargeField rstEmployees2!Photo, !Photo 
     .Update 
     
     ' Delete new record because this is a demonstration. 
     .Bookmark = .LastModified 
     .Delete 
     .Close 
     End With 
     
     rstEmployees2.Close 
     dbsNorthwind.Close 
     
    End Sub 
     
    Function CopyLargeField(fldSource As Field, _ 
     fldDestination As Field) 
     
     ' Set size of chunk in bytes. 
     Const conChunkSize = 32768 
     
     Dim lngOffset As Long 
     Dim lngTotalSize As Long 
     Dim strChunk As String 
     
     ' Copy the photo from one Recordset to the other in 32K 
     ' chunks until the entire field is copied. 
     lngTotalSize = fldSource.FieldSize 
     Do While lngOffset < lngTotalSize 
     strChunk = fldSource.GetChunk(lngOffset, conChunkSize) 
     fldDestination.AppendChunk strChunk 
     lngOffset = lngOffset + conChunkSize 
     Loop 
     
    End Function
```