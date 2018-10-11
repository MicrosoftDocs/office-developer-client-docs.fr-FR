---
title: Recordset.PercentPosition Property (DAO)
TOCTitle: PercentPosition Property
ms:assetid: aebbda44-ed72-7a6c-0cd5-28c8997d4d96
ms:mtpsurl: https://msdn.microsoft.com/library/Ff821751(v=office.15)
ms:contentKeyID: 48547077
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 5793d2b73e424726e91cf3bbb54b09db59123b92
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470656"
---
# <a name="recordsetpercentposition-property-dao"></a><span data-ttu-id="f2363-102">Recordset.PercentPosition Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="f2363-102">Recordset.PercentPosition Property (DAO)</span></span>


<span data-ttu-id="f2363-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="f2363-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="f2363-104">Définit ou renvoie une valeur indiquant l'emplacement approximatif de l'enregistrement actif dans l'objet **[Recordset](recordset-object-dao.md)** en fonction du pourcentage d'enregistrements dans cet objet **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="f2363-104">Sets or returns a value indicating the approximate location of the current record in the **[Recordset](recordset-object-dao.md)** object based on a percentage of the records in the **Recordset**.</span></span>

## <a name="syntax"></a><span data-ttu-id="f2363-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="f2363-105">Syntax</span></span>

<span data-ttu-id="f2363-106">*expression* . PercentPosition</span><span class="sxs-lookup"><span data-stu-id="f2363-106">*expression* .PercentPosition</span></span>

<span data-ttu-id="f2363-107">*expression* Variable qui représente un objet **Recordset** .</span><span class="sxs-lookup"><span data-stu-id="f2363-107">*expression* A variable that represents a **Recordset** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="f2363-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="f2363-108">Remarks</span></span>

<span data-ttu-id="f2363-p101">Pour indiquer ou modifier la position approximative de l'enregistrement actif dans un objet **Recordset**, vous pouvez vérifier ou définir la propriété **PercentPosition**. Lorsque vous manipulez un objet **Recordset** de type feuille de réponse dynamique ou instantané, ouvert directement à partir d'une table de base, commencez par remplir l'objet **Recordset** en accédant au dernier enregistrement avant de définir ou vérifier la propriété **PercentPosition**. Si vous utilisez la propriété **PercentPosition** avant d'avoir entièrement rempli l'objet **Recordset**, le volume de déplacements est relatif au nombre d'enregistrements accédés, tel qu'il est spécifié dans le paramètre de la propriété **[RecordCount](recordset-recordcount-property-dao.md)**. Vous pouvez accéder au dernier enregistrement à l'aide de la méthode **[MoveLast](recordset-movelast-method-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="f2363-p101">To indicate or change the approximate position of the current record in a **Recordset** object, you can check or set the **PercentPosition** property. When working with a dynaset- or snapshot-type **Recordset** object opened directly from a base table, first populate the **Recordset** object by moving to the last record before you set or check the **PercentPosition** property. If you use the **PercentPosition** property before fully populating the **Recordset** object, the amount of movement is relative to the number of records accessed as indicated by the **[RecordCount](recordset-recordcount-property-dao.md)** property setting. You can move to the last record by using the **[MoveLast](recordset-movelast-method-dao.md)** method.</span></span>


> [!NOTE]
> <P><span data-ttu-id="f2363-113">[!REMARQUE] Il est déconseillé d'utiliser la propriété <STRONG>PercentPosition</STRONG> si votre intention est de faire d'un enregistrement spécifique de l'objet <STRONG>Recordset</STRONG> l'enregistrement actif. La propriété <STRONG><A href="recordset-bookmark-property-dao.md">Bookmark</A></STRONG> est plus adaptée pour cela.</span><span class="sxs-lookup"><span data-stu-id="f2363-113">Using the <STRONG>PercentPosition</STRONG> property to move the current record to a specific record in a <STRONG>Recordset</STRONG> object isn't recommended?the <STRONG><A href="recordset-bookmark-property-dao.md">Bookmark</A></STRONG> property is better suited for this task.</span></span></P>



<span data-ttu-id="f2363-p102">Dès que vous définissez une valeur pour la propriété **PercentPosition**, l'enregistrement qui occupe la position approximative correspondant à cette valeur devient l'enregistrement actif et la propriété **PercentPosition** est réinitialisée avec une valeur qui reflète la position approximative de l'enregistrement actif. Par exemple, si votre objet **Recordset** contient seulement cinq enregistrements et que vous définissez la valeur 77 pour la propriété **PercentPosition**, la valeur renvoyée par la propriété **PercentPosition** sera vraisemblablement 80 et non 77.</span><span class="sxs-lookup"><span data-stu-id="f2363-p102">Once you set the **PercentPosition** property to a value, the record at the approximate position corresponding to that value becomes current, and the **PercentPosition** property is reset to a value that reflects the approximate position of the current record. For example, if your **Recordset** object contains only five records, and you set its **PercentPosition** property value to 77, the value returned from the **PercentPosition** property may be 80, not 77.</span></span>

<span data-ttu-id="f2363-116">La propriété **PercentPosition** s’applique à tous les types d’objets **Recordset** à l’exception des objets **Recordset** de type avant uniquement ou les objets **Recordset** ouverts à partir de requêtes directes sur des bases de données distantes.</span><span class="sxs-lookup"><span data-stu-id="f2363-116">The **PercentPosition** property applies to all types of **Recordset** objects except for forward–only–type **Recordset** objects or **Recordset** objects opened from pass-through queries against remote databases.</span></span>

<span data-ttu-id="f2363-117">Vous pouvez utiliser la propriété **PercentPosition** avec une barre de défilement dans un formulaire ou une zone de texte pour indiquer l'emplacement de l'enregistrement actif dans un objet **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="f2363-117">You can use the **PercentPosition** property with a scroll bar on a form or text box to indicate the location of the current record in a **Recordset** object.</span></span>

## <a name="example"></a><span data-ttu-id="f2363-118">Exemple</span><span class="sxs-lookup"><span data-stu-id="f2363-118">Example</span></span>

<span data-ttu-id="f2363-119">Cet exemple utilise la propriété **PercentPosition** pour indiquer la position du pointeur d'enregistrement actif par rapport au début de l'objet **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="f2363-119">This example uses the **PercentPosition** property to show the position of the current record pointer relative to the beginning of the **Recordset**.</span></span>

```vb
    Sub PercentPositionX() 
     
     Dim dbsNorthwind As Database 
     Dim rstProducts As Recordset 
     Dim strFind As String 
     Dim strMessage As String 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     ' PercentPosition only works with dynasets or snapshots. 
     Set rstProducts = dbsNorthwind.OpenRecordset( _ 
     "SELECT ProductName FROM Products " & _ 
     "ORDER BY ProductName", dbOpenSnapshot) 
     
     With rstProducts 
     ' Populate the Recordset. 
     .MoveLast 
     .MoveFirst 
     
     Do While True 
     ' Show current record information and ask user 
     ' for input. 
     strMessage = "Product: " & !ProductName & vbCr & _ 
     "The record pointer is " & _ 
     Format(.PercentPosition, "##0.0") & _ 
     "% from the " & vbCr & _ 
     "beginning of the Recordset." & vbCr & _ 
     "Please enter a character search string " & _ 
     "for a product name." 
     strFind = Trim(InputBox(strMessage)) 
     If strFind = "" Then Exit Do 
     
     ' Try to find a record matching the search string. 
     .FindFirst "ProductName >= '" & strFind & "'" 
     If .NoMatch Then .MoveLast 
     Loop 
     
     .Close 
     End With 
     
     dbsNorthwind.Close 
     
    End Sub
```