---
title: Index.IgnoreNulls Property (DAO)
TOCTitle: IgnoreNulls Property
ms:assetid: f49f17b8-d7c1-18ab-07a8-e1be61488519
ms:mtpsurl: https://msdn.microsoft.com/library/Ff836698(v=office.15)
ms:contentKeyID: 48548694
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052931
f1_categories:
- Office.Version=v15
ms.openlocfilehash: b6c190b7a61e26ff6e4abedc1a19bde26a1de426
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470948"
---
# <a name="indexignorenulls-property-dao"></a><span data-ttu-id="65766-102">Index.IgnoreNulls Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="65766-102">Index.IgnoreNulls Property (DAO)</span></span>


<span data-ttu-id="65766-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="65766-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="65766-104">Définit ou renvoie une valeur qui indique si les enregistrements ayant des valeurs nulles dans leurs champs d'index ont des entrées d'index (espaces de travail Microsoft Access uniquement).</span><span class="sxs-lookup"><span data-stu-id="65766-104">Sets or returns a value that indicates whether records that have Null values in their index fields have index entries (Microsoft Access workspaces only).</span></span>

## <a name="syntax"></a><span data-ttu-id="65766-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="65766-105">Syntax</span></span>

<span data-ttu-id="65766-106">*expression* . IgnoreNulls</span><span class="sxs-lookup"><span data-stu-id="65766-106">*expression* .IgnoreNulls</span></span>

<span data-ttu-id="65766-107">*expression* Variable qui représente un objet **Index** .</span><span class="sxs-lookup"><span data-stu-id="65766-107">*expression* A variable that represents an **Index** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="65766-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="65766-108">Remarks</span></span>

<span data-ttu-id="65766-109">Cette propriété est en lecture/écriture pour un nouvel objet **[Index](index-object-dao.md)** pas encore ajouté à une collection et en lecture seule pour un objet **Index** existant de la collection **[Indexes](indexes-collection-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="65766-109">This property is read/write for a new **[Index](index-object-dao.md)** object not yet appended to a collection and read-only for an existing **Index** object in an **[Indexes](indexes-collection-dao.md)** collection.</span></span>

<span data-ttu-id="65766-p101">Pour accélérer le processus de recherche des enregistrements, vous pouvez définir un index pour un champ. Si vous autorisez les entrées de type **null** dans un champ indexé et prévoyez d'avoir beaucoup d'entrées **null**, vous pouvez définir la propriété **IgnoreNulls** de l'objet **Index** sur **True** pour réduire la quantité d'espace de stockage utilisé par l'index.</span><span class="sxs-lookup"><span data-stu-id="65766-p101">To speed up the process of searching for records, you can define an index for a field. If you allow **null** entries in an indexed field and expect many of the entries to be **null**, you can set the **IgnoreNulls** property for the **Index** object to **True** to reduce the amount of storage space that the index uses.</span></span>

<span data-ttu-id="65766-112">Les paramètres de propriété **IgnoreNulls** et **[Required](field-required-property-dao.md)** déterminent ensemble si un enregistrement doté d'une valeur d'index de type **null** a une entrée d'index.</span><span class="sxs-lookup"><span data-stu-id="65766-112">The **IgnoreNulls** property setting and the **[Required](field-required-property-dao.md)** property setting together determine whether a record with a **null** index value has an index entry.</span></span>

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="65766-113">Si IgnoreNulls est</span><span class="sxs-lookup"><span data-stu-id="65766-113">If IgnoreNulls is</span></span></p></th>
<th><p><span data-ttu-id="65766-114">Et Required est</span><span class="sxs-lookup"><span data-stu-id="65766-114">And Required is</span></span></p></th>
<th><p><span data-ttu-id="65766-115">Alors</span><span class="sxs-lookup"><span data-stu-id="65766-115">Then</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="65766-116">True</span><span class="sxs-lookup"><span data-stu-id="65766-116">True</span></span></p></td>
<td><p><span data-ttu-id="65766-117">False</span><span class="sxs-lookup"><span data-stu-id="65766-117">False</span></span></p></td>
<td><p><span data-ttu-id="65766-118">Une valeur nulle est autorisée dans le champ d'index ; aucune entrée d'index n'est ajoutée.</span><span class="sxs-lookup"><span data-stu-id="65766-118">A null value is allowed in the index field; no index entry added.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="65766-119">False</span><span class="sxs-lookup"><span data-stu-id="65766-119">False</span></span></p></td>
<td><p><span data-ttu-id="65766-120">False</span><span class="sxs-lookup"><span data-stu-id="65766-120">False</span></span></p></td>
<td><p><span data-ttu-id="65766-121">Une valeur nulle est autorisée dans le champ d'index ; entrée d'index ajoutée.</span><span class="sxs-lookup"><span data-stu-id="65766-121">A null value is allowed in the index field; index entry added.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="65766-122">True ou False</span><span class="sxs-lookup"><span data-stu-id="65766-122">True or False</span></span></p></td>
<td><p><span data-ttu-id="65766-123">True</span><span class="sxs-lookup"><span data-stu-id="65766-123">True</span></span></p></td>
<td><p><span data-ttu-id="65766-124">Aucune valeur nulle n'est autorisée dans le champ d'index ; aucune entrée d'index n'est ajoutée.</span><span class="sxs-lookup"><span data-stu-id="65766-124">A null value isn't allowed in the index field; no index entry added.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="example"></a><span data-ttu-id="65766-125">Exemple</span><span class="sxs-lookup"><span data-stu-id="65766-125">Example</span></span>

<span data-ttu-id="65766-126">Cet exemple définit la propriété **IgnoreNulls** d'un nouvel **Index** sur **True** ou **False** en fonction de l'entrée de l'utilisateur, puis en démontre l'effet sur un **Recordset** contenant un enregistrement dont le champ de clé contient une valeur **Null**.</span><span class="sxs-lookup"><span data-stu-id="65766-126">This example sets the **IgnoreNulls** property of a new **Index** to **True** or **False** based on user input, and then demonstrates the effect on a **Recordset** with a record whose key field contains a **Null** value.</span></span>

```vb
    Sub IgnoreNullsX() 
     
     Dim dbsNorthwind As Database 
     Dim tdfEmployees As TableDef 
     Dim idxNew As Index 
     Dim rstEmployees As Recordset 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     Set tdfEmployees = dbsNorthwind!Employees 
     
     With tdfEmployees 
     ' Create a new Index object. 
     Set idxNew = .CreateIndex("NewIndex") 
     idxNew.Fields.Append idxNew.CreateField("Country") 
     
     ' Set the IgnoreNulls property of the new Index object 
     ' based on the user's input. 
     Select Case MsgBox("Set IgnoreNulls to True?", _ 
     vbYesNoCancel) 
     Case vbYes 
     idxNew.IgnoreNulls = True 
     Case vbNo 
     idxNew.IgnoreNulls = False 
     Case Else 
     dbsNorthwind.Close 
     End 
     End Select 
     
     ' Append the new Index object to the Indexes 
     ' collection of the Employees table. 
     .Indexes.Append idxNew 
     .Indexes.Refresh 
     End With 
     
     Set rstEmployees = _ 
     dbsNorthwind.OpenRecordset("Employees") 
     
     With rstEmployees 
     ' Add a new record to the Employees table. 
     .AddNew 
     !FirstName = "Gary" 
     !LastName = "Haarsager" 
     .Update 
     
     ' Use the new index to set the order of the records. 
     .Index = idxNew.Name 
     .MoveFirst 
     
     Debug.Print "Index = " & .Index & _ 
     ", IgnoreNulls = " & idxNew.IgnoreNulls 
     Debug.Print " Country - Name" 
     
     ' Enumerate the Recordset. The value of the 
     ' IgnoreNulls property will determine if the newly 
     ' added record appears in the output. 
     Do While Not .EOF 
     Debug.Print " " & _ 
     IIf(IsNull(!Country), "[Null]", !Country) & _ 
     " - " & !FirstName & " " & !LastName 
     .MoveNext 
     Loop 
     
     ' Delete new record because this is a demonstration. 
     .Index = "" 
     .Bookmark = .LastModified 
     .Delete 
     .Close 
     End With 
     
     ' Delete new Index because this is a demonstration. 
     tdfEmployees.Indexes.Delete idxNew.Name 
     dbsNorthwind.Close 
     
    End Sub
```