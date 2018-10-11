---
title: Field2.OrdinalPosition Property (DAO)
TOCTitle: OrdinalPosition Property
ms:assetid: 55d89611-ad07-990d-fc33-f81d59472430
ms:mtpsurl: https://msdn.microsoft.com/library/Ff194179(v=office.15)
ms:contentKeyID: 48544937
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052899
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 5c086263ffe2267a8f01cfae2e67db125313c0b2
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469766"
---
# <a name="field2ordinalposition-property-dao"></a><span data-ttu-id="7a0d5-102">Field2.OrdinalPosition Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="7a0d5-102">Field2.OrdinalPosition Property (DAO)</span></span>


<span data-ttu-id="7a0d5-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="7a0d5-103">**Applies to**: Access 2013 | Office 2013</span></span>


<span data-ttu-id="7a0d5-p101">Définit ou renvoie la position relative d'un objet **Field2** au sein d'une collection **[Fields](fields-collection-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-p101">Sets or returns the relative position of a **Field2** object within a **[Fields](fields-collection-dao.md)** collection. .</span></span>

## <a name="syntax"></a><span data-ttu-id="7a0d5-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="7a0d5-106">Syntax</span></span>

<span data-ttu-id="7a0d5-107">*expression* . OrdinalPosition</span><span class="sxs-lookup"><span data-stu-id="7a0d5-107">*expression* .OrdinalPosition</span></span>

<span data-ttu-id="7a0d5-108">*expression* Variable qui représente un objet **Field2** .</span><span class="sxs-lookup"><span data-stu-id="7a0d5-108">*expression* A variable that represents a **Field2** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="7a0d5-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="7a0d5-109">Remarks</span></span>

<span data-ttu-id="7a0d5-110">Pour un objet pas encore ajouté à la collection **Fields**, cette propriété est en lecture/écriture.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-110">For an object not yet appended to the **Fields** collection, this property is read/write.</span></span>

<span data-ttu-id="7a0d5-111">La valeur par défaut est 0.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-111">The default is 0.</span></span>

<span data-ttu-id="7a0d5-112">La disponibilité de la propriété **OrdinalPosition** dépend de l'objet contenant la collection **Fields**, comme illustré dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-112">The availability of the **OrdinalPosition** property depends on the object that contains the **Fields** collection, as shown in the following table.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="7a0d5-113">Si la collection Fields appartient à un</span><span class="sxs-lookup"><span data-stu-id="7a0d5-113">If the Fields collection belongs to a</span></span></p></th>
<th><p><span data-ttu-id="7a0d5-114">
Alors OrdinalPosition est</span><span class="sxs-lookup"><span data-stu-id="7a0d5-114">Then OrdinalPosition is</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="7a0d5-115">							objet <strong>Index</strong></span><span class="sxs-lookup"><span data-stu-id="7a0d5-115"><strong>Index</strong> object</span></span></p></td>
<td><p><span data-ttu-id="7a0d5-116">Non pris en charge</span><span class="sxs-lookup"><span data-stu-id="7a0d5-116">Not supported</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="7a0d5-117">							objet <strong>QueryDef</strong></span><span class="sxs-lookup"><span data-stu-id="7a0d5-117"><strong>QueryDef</strong> object</span></span></p></td>
<td><p><span data-ttu-id="7a0d5-118">Lecture seule</span><span class="sxs-lookup"><span data-stu-id="7a0d5-118">Read-only</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="7a0d5-119">							objet <strong>Recordset</strong></span><span class="sxs-lookup"><span data-stu-id="7a0d5-119"><strong>Recordset</strong> object</span></span></p></td>
<td><p><span data-ttu-id="7a0d5-120">Lecture seule</span><span class="sxs-lookup"><span data-stu-id="7a0d5-120">Read-only</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="7a0d5-121">							objet <strong>Relation</strong></span><span class="sxs-lookup"><span data-stu-id="7a0d5-121"><strong>Relation</strong> object</span></span></p></td>
<td><p><span data-ttu-id="7a0d5-122">Non pris en charge</span><span class="sxs-lookup"><span data-stu-id="7a0d5-122">Not supported</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="7a0d5-123">							objet <strong>TableDef</strong></span><span class="sxs-lookup"><span data-stu-id="7a0d5-123"><strong>TableDef</strong> object</span></span></p></td>
<td><p><span data-ttu-id="7a0d5-124">En lecture-écriture.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-124">Read/write</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="7a0d5-p102">En règle générale, la position ordinale d’un objet ajouté à une collection dépend de l’ordre dans lequel vous ajoutez l’objet. Le premier objet est ajouté a la première position (0), le second à la seconde position (1) et ainsi de suite. Le dernier objet ajouté est à la position ordinale compte – 1, où compte est le nombre d’objets de la collection tel que spécifié par le paramètre de propriété **[Count](containers-count-property-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-p102">Generally, the ordinal position of an object that you append to a collection depends on the order in which you append the object. The first appended object is in the first position (0), the second appended object is in the second position (1), and so on. The last appended object is in ordinal position count – 1, where count is the number of objects in the collection as specified by the **[Count](containers-count-property-dao.md)** property setting.</span></span>

<span data-ttu-id="7a0d5-128">La propriété **OrdinalPosition** vous permet de spécifier une position ordinale pour les nouveaux objets **Field2** qui diffère de l'ordre dans lequel vous les ajoutez à une collection.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-128">You can use the **OrdinalPosition** property to specify an ordinal position for new **Field2** objects that differs from the order in which you append those objects to a collection.</span></span> <span data-ttu-id="7a0d5-129">Ceci vous permet de spécifier un ordre de champs pour vos tables, requêtes et jeux d'enregistrements lorsque vous les utilisez dans une application.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-129">This enables you to specify a field order for your tables, queries, and recordsets when you use them in an application.</span></span> <span data-ttu-id="7a0d5-130">Par exemple, l’ordre dans lequel les champs sont retournés dans une instruction SELECT \* requête est déterminée par les valeurs actuelles de la propriété **OrdinalPosition** .</span><span class="sxs-lookup"><span data-stu-id="7a0d5-130">For example, the order in which fields are returned in a SELECT \* query is determined by the current **OrdinalPosition** property values.</span></span>

<span data-ttu-id="7a0d5-131">Vous pouvez à tout moment redéfinir l'ordre de renvoi des champs dans les jeux d'enregistrements en définissant la propriété **OrdinalPosition** sur tout entier positif.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-131">You can permanently reset the order in which fields are returned in recordsets by setting the **OrdinalPosition** property to any positive integer.</span></span>

<span data-ttu-id="7a0d5-p104">Deux champs **Field2** ou plus peuvent avoir la même valeur de propriété **OrdinalPosition**, auquel cas ils sont classés par ordre alphabétique. Par exemple, si vous avez un champ appelé Age défini sur 4 et un second champ appelé Poids sur 4, Poids est renvoyé après Age.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-p104">Two or more **Field2** objects in the same collection can have the same **OrdinalPosition** property value, in which case they will be ordered alphabetically. For example, if you have a field named Age set to 4 and you set a second field named Weight to 4, Weight is returned after Age.</span></span>

<span data-ttu-id="7a0d5-p105">Vous pouvez spécifier un nombre supérieur au nombre de champs moins 1. Le champ est renvoyé dans un ordre relatif au nombre le plus élevé. Par exemple, si vous définissez la propriété **OrdinalPosition** d'un champ sur 20 (et qu'il n'y a que 5 champs) et si vous avez défini la propriété **OrdinalPosition** des deux autres champs sur 10 et 30, respectivement, le champ défini sur 20 est renvoyé entre les champs définis sur 10 et 30.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-p105">You can specify a number that is greater than the number of fields minus 1. The field will be returned in an order relative to the largest number. For example, if you set a field's **OrdinalPosition** property to 20 (and there are only 5 fields) and you've set the **OrdinalPosition** property for two other fields to 10 and 30, respectively, the field set to 20 is returned between the fields set to 10 and 30.</span></span>


> [!NOTE]
> <P><span data-ttu-id="7a0d5-p106">[!REMARQUE] Même si la collection <STRONG>Fields</STRONG> d'un objet <STRONG><A href="tabledef-object-dao.md">TableDef</A></STRONG> n'a pas été rafraîchie, l'ordre des champs dans un objet <STRONG><A href="recordset-object-dao.md">Recordset</A></STRONG> ouvert à partir de l'objet <STRONG>TableDef</STRONG> reflète les données <STRONG>OrdinalPosition</STRONG> de l'objet <STRONG>TableDef</STRONG>. Un objet <STRONG>Recordset</STRONG> de type table a les mêmes données <STRONG>OrdinalPosition</STRONG> en tant que table sous-jacente, mais tout autre type d'objet <STRONG>Recordset</STRONG> a des nouvelles données <STRONG>OrdinalPosition</STRONG> (commençant par 0) qui suivent l'ordre déterminé par les données <STRONG>OrdinalPosition</STRONG> de l'objet <STRONG>TableDef</STRONG>.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-p106">Even if the <STRONG>Fields</STRONG> collection of a <STRONG><A href="tabledef-object-dao.md">TableDef</A></STRONG> has not been refreshed, the field order in a <STRONG><A href="recordset-object-dao.md">Recordset</A></STRONG> opened from the <STRONG>TableDef</STRONG> will reflect the <STRONG>OrdinalPosition</STRONG> data of the <STRONG>TableDef</STRONG> object. A table-type <STRONG>Recordset</STRONG> will have the same <STRONG>OrdinalPosition</STRONG> data as the underlying table, but any other type of <STRONG>Recordset</STRONG> will have new <STRONG>OrdinalPosition</STRONG> data (starting with 0) that follow the order determined by the <STRONG>OrdinalPosition</STRONG> data of the <STRONG>TableDef</STRONG>.</span></span></P>



## <a name="example"></a><span data-ttu-id="7a0d5-139">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a0d5-139">Example</span></span>

<span data-ttu-id="7a0d5-p107">Cet exemple modifie les valeurs de propriété **OrdinalPosition** de la **TableDef** Employees afin de contrôler l'ordre des **Field2** du **Recordset** résultant. En définissant la **OrdinalPosition** de tous les **Fields** sur 1, tous les **Fields** des **Recordset** résultants sont classés par ordre alphabétique. Notez que les valeurs **OrdinalPosition** du **Recordset** ne correspondent pas à celles de la **TableDef**, mais reflètent simplement le résultat final des modifications de **TableDef**.</span><span class="sxs-lookup"><span data-stu-id="7a0d5-p107">This example changes the **OrdinalPosition** property values in the Employees **TableDef** in order to control the **Field2** order in a resulting **Recordset**. By setting the **OrdinalPosition** of all the **Fields** to 1, any resulting **Recordset** will order the **Fields** alphabetically. Note that the **OrdinalPosition** values in the **Recordset** don't match the values in the **TableDef**, but simply reflect the end result of the **TableDef** changes.</span></span>

```vb
    Sub OrdinalPositionX() 
     
     Dim dbsNorthwind As Database 
     Dim tdfEmployees As TableDef 
     Dim aintPosition() As Integer 
     Dim astrFieldName() As String 
     Dim intTemp As Integer 
     Dim fldTemp As Field2 
     Dim rstEmployees As Recordset 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     Set tdfEmployees = dbsNorthwind.TableDefs("Employees") 
     
     With tdfEmployees 
     ' Display and store original OrdinalPosition data. 
     Debug.Print _ 
     "Original OrdinalPosition data in TableDef." 
     ReDim aintPosition(0 To .Fields.Count - 1) As Integer 
     ReDim astrFieldName(0 To .Fields.Count - 1) As String 
     For intTemp = 0 To .Fields.Count - 1 
     aintPosition(intTemp) = _ 
     .Fields(intTemp).OrdinalPosition 
     astrFieldName(intTemp) = .Fields(intTemp).Name 
     Debug.Print , aintPosition(intTemp), _ 
     astrFieldName(intTemp) 
     Next intTemp 
     
     ' Change OrdinalPosition data. 
     For Each fldTemp In .Fields 
     fldTemp.OrdinalPosition = 1 
     Next fldTemp 
     
     ' Open new Recordset object to show how the 
     ' OrdinalPosition data has affected the record order. 
     Debug.Print _ 
     "OrdinalPosition data from resulting Recordset." 
     Set rstEmployees = dbsNorthwind.OpenRecordset( _ 
     "SELECT * FROM Employees") 
     For Each fldTemp In rstEmployees.Fields 
     Debug.Print , fldTemp.OrdinalPosition, fldTemp.Name 
     Next fldTemp 
     rstEmployees.Close 
     
     ' Restore original OrdinalPosition data because this is 
     ' a demonstration. 
     For intTemp = 0 To .Fields.Count - 1 
     .Fields(astrFieldName(intTemp)).OrdinalPosition = _ 
     aintPosition(intTemp) 
     Next intTemp 
     
     End With 
     
     dbsNorthwind.Close 
     
    End Sub
```