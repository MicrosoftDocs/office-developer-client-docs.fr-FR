---
title: Recordset2.AbsolutePosition Property (DAO)
TOCTitle: AbsolutePosition Property
ms:assetid: 91ca203f-0c80-67f4-e180-415b6af05030
ms:mtpsurl: https://msdn.microsoft.com/library/Ff197637(v=office.15)
ms:contentKeyID: 48546358
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1053074
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 4f521daab352e090c52fd7a94bdd28abb0e34291
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471471"
---
# <a name="recordset2absoluteposition-property-dao"></a><span data-ttu-id="1507d-102">Recordset2.AbsolutePosition Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="1507d-102">Recordset2.AbsolutePosition Property (DAO)</span></span>


<span data-ttu-id="1507d-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="1507d-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="1507d-104">Définit ou renvoie le nombre d’enregistrements relatif de l’enregistrement actif d’un objet **Recordset2**.</span><span class="sxs-lookup"><span data-stu-id="1507d-104">Sets or returns the relative record number of a **Recordset2** object's current record.</span></span>

## <a name="syntax"></a><span data-ttu-id="1507d-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="1507d-105">Syntax</span></span>

<span data-ttu-id="1507d-106">*expression* . AbsolutePosition</span><span class="sxs-lookup"><span data-stu-id="1507d-106">*expression* .AbsolutePosition</span></span>

<span data-ttu-id="1507d-107">*expression* Variable qui représente un objet **Recordset2** .</span><span class="sxs-lookup"><span data-stu-id="1507d-107">*expression* A variable that represents a **Recordset2** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="1507d-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="1507d-108">Remarks</span></span>

<span data-ttu-id="1507d-p101">Vous pouvez utiliser la propriété **AbsolutePosition** pour placer le pointeur d'enregistrement actif sur un enregistrement spécifique en fonction de sa position ordinale dans un objet **Recordset2** de type instantané ou feuille de réponse dynamique. Vous pouvez également déterminer le numéro de l'enregistrement actif en vérifiant le paramètre de la propriété **AbsolutePosition**.</span><span class="sxs-lookup"><span data-stu-id="1507d-p101">You can use the **AbsolutePosition** property to position the current record pointer to a specific record based on its ordinal position in a dynaset- or snapshot-type **Recordset2** object. You can also determine the current record number by checking the **AbsolutePosition** property setting.</span></span>

<span data-ttu-id="1507d-p102">Dans la mesure où la valeur de la propriété **AbsolutePosition** est en base zéro (c'est-à-dire que la valeur 0 référence le premier enregistrement de l'objet **Recordset2**), vous ne pouvez pas lui affecter une valeur supérieure ou égale au nombre d'enregistrements remplis. Si c'est le cas, une erreur piégeable se produit. Vous pouvez déterminer le nombre d'enregistrements remplis dans l'objet **Recordset2** en vérifiant la valeur de la propriété **RecordCount**. La valeur maximale autorisée de la propriété **AbsolutePosition** correspond à la valeur de la propriété **RecordCount** moins 1.</span><span class="sxs-lookup"><span data-stu-id="1507d-p102">Because the **AbsolutePosition** property value is zero-based (that is, a setting of 0 refers to the first record in the **Recordset2** object), you cannot set it to a value greater than or equal to the number of populated records; doing so causes a trappable error. You can determine the number of populated records in the **Recordset2** object by checking the **RecordCount** property setting. The maximum allowable setting for the **AbsolutePosition** property is the value of the **RecordCount** property minus 1.</span></span>

<span data-ttu-id="1507d-114">Si aucun enregistrement actif, en tant que lorsqu’il n’y a aucun enregistrement dans l’objet **Recordset2** , **AbsolutePosition** renvoie – 1.</span><span class="sxs-lookup"><span data-stu-id="1507d-114">If there is no current record, as when there are no records in the **Recordset2** object, **AbsolutePosition** returns –1.</span></span> <span data-ttu-id="1507d-115">Si l'enregistrement actif est supprimé, la valeur de la propriété **AbsolutePosition** n'est pas définie et une erreur interceptable se produit s'il est référencé.</span><span class="sxs-lookup"><span data-stu-id="1507d-115">If the current record is deleted, the **AbsolutePosition** property value isn't defined, and a trappable error occurs if it's referenced.</span></span> <span data-ttu-id="1507d-116">Les nouveaux enregistrements sont ajoutés à la fin de la séquence.</span><span class="sxs-lookup"><span data-stu-id="1507d-116">New records are added to the end of the sequence.</span></span>

<span data-ttu-id="1507d-p104">Vous ne devez pas utiliser cette propriété comme numéro d'enregistrement de substitution. La meilleure façon de conserver et revenir à une position donnée consiste à utiliser les signets. Ils constituent la seule possibilité de positionner l'enregistrement actif dans tous les types d'objets **Recordset2**. La position d'un enregistrement varie lorsqu'un ou plusieurs enregistrements qui le précèdent sont supprimés. Il n'existe aucune assurance qu'un enregistrement conservera la même position absolue si l'objet **Recordset2** est recréé car l'ordre des enregistrements individuels dans un objet **Recordset** n'est jamais garanti sauf s'il est créé avec une instruction SQL comportant une clause ORDER BY.</span><span class="sxs-lookup"><span data-stu-id="1507d-p104">You shouldn't use this property as a surrogate record number. Bookmarks are still the recommended way of retaining and returning to a given position and are the only way to position the current record across all types of **Recordset2** objects. In particular, the position of a record changes when one or more records preceding it are deleted. There is also no assurance that a record will have the same absolute position if the **Recordset2** object is re-created again because the order of individual records within a **Recordset** object isn't guaranteed unless it's created with an SQL statement by using an ORDER BY clause.</span></span>


> [!NOTE]
> <UL>
> <LI>
> <P><span data-ttu-id="1507d-p105">Si vous affectez une valeur supérieure à zéro à la propriété <STRONG>AbsolutePosition</STRONG> d'un objet <STRONG>Recordset2</STRONG> récemment ouvert mais toujours vide, une erreur piégeable se produit. Remplissez d'abord l'objet <STRONG>Recordset2</STRONG> avec la méthode <STRONG>MoveLast</STRONG>.</span><span class="sxs-lookup"><span data-stu-id="1507d-p105">Setting the <STRONG>AbsolutePosition</STRONG> property to a value greater than zero on a newly opened but unpopulated <STRONG>Recordset2</STRONG> object causes a trappable error. Populate the <STRONG>Recordset2</STRONG> object first with the <STRONG>MoveLast</STRONG> method.</span></span></P>
> <LI>
> <P><span data-ttu-id="1507d-123">La propriété <STRONG>AbsolutePosition</STRONG> n’est pas disponible sur les objets <STRONG>Recordset2</STRONG> de type avant uniquement ou les objets <STRONG>Recordset2</STRONG> ouverts à partir de requêtes directes sur des bases de données ODBC connectées au moteur de base de données Microsoft Access.</span><span class="sxs-lookup"><span data-stu-id="1507d-123">The <STRONG>AbsolutePosition</STRONG> property isn't available on forward–only–type <STRONG>Recordset2</STRONG> objects, or on <STRONG>Recordset2</STRONG> objects opened from pass-through queries against Microsoft Access database engine-connected ODBC databases.</span></span></P></LI></UL>



## <a name="example"></a><span data-ttu-id="1507d-124">Exemple</span><span class="sxs-lookup"><span data-stu-id="1507d-124">Example</span></span>

<span data-ttu-id="1507d-125">Cet exemple utilise la propriété **AbsolutePosition** pour suivre la progression d'une boucle qui énumère tous les enregistrements d'un objet **Recordset2**.</span><span class="sxs-lookup"><span data-stu-id="1507d-125">This example uses the **AbsolutePosition** property to track the progress of a loop that enumerates all the records of a **Recordset2**.</span></span>

```vb
    Sub AbsolutePositionX() 
     
       Dim dbsNorthwind As Database 
       Dim rstEmployees As Recordset2 
       Dim strMessage As String 
     
       Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
       ' AbsolutePosition only works with dynasets or snapshots. 
       Set rstEmployees = _ 
          dbsNorthwind.OpenRecordset("Employees", _ 
          dbOpenSnapshot) 
     
       With rstEmployees 
          ' Populate Recordset. 
          .MoveLast 
          .MoveFirst 
     
          ' Enumerate Recordset. 
          Do While Not .EOF 
             ' Display current record information. Add 1 to  
             ' AbsolutePosition value because it is zero-based. 
             strMessage = "Employee: " & !LastName & vbCr & _ 
                "(record " & (.AbsolutePosition + 1) & _ 
                " of " & .RecordCount & ")" 
             If MsgBox(strMessage, vbOKCancel) = vbCancel _ 
                Then Exit Do 
             .MoveNext 
          Loop 
     
          .Close 
       End With 
     
       dbsNorthwind.Close 
     
    End Sub
```