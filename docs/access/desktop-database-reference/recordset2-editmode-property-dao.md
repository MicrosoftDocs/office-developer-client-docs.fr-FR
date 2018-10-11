---
title: Recordset2.EditMode Property (DAO)
TOCTitle: EditMode Property
ms:assetid: fd61ea2b-e7d7-195f-4114-87e54eba2451
ms:mtpsurl: https://msdn.microsoft.com/library/Ff837240(v=office.15)
ms:contentKeyID: 48548914
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1053080
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 7b8ff4fa104dd153faf6eb1a50d2e922e5e5021d
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469296"
---
# <a name="recordset2editmode-property-dao"></a><span data-ttu-id="38d03-102">Recordset2.EditMode Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="38d03-102">Recordset2.EditMode Property (DAO)</span></span>


<span data-ttu-id="38d03-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="38d03-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="38d03-104">Renvoie une valeur qui indique l'état de modification de l'enregistrement actif.</span><span class="sxs-lookup"><span data-stu-id="38d03-104">Returns a value that indicates the state of editing for the current record.</span></span>

## <a name="syntax"></a><span data-ttu-id="38d03-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="38d03-105">Syntax</span></span>

<span data-ttu-id="38d03-106">*expression* . EditMode</span><span class="sxs-lookup"><span data-stu-id="38d03-106">*expression* .EditMode</span></span>

<span data-ttu-id="38d03-107">*expression* Variable qui représente un objet **Recordset2** .</span><span class="sxs-lookup"><span data-stu-id="38d03-107">*expression* A variable that represents a **Recordset2** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="38d03-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="38d03-108">Remarks</span></span>

<span data-ttu-id="38d03-p101">La valeur de retour est une donnée de type **Long** qui indique l'état de modification. La valeur peut être une des constantes **[EditModeEnum](editmodeenum-enumeration-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="38d03-p101">The return value is a **Long** that indicates the state of editing. The value can be one of the **[EditModeEnum](editmodeenum-enumeration-dao.md)** constants.</span></span>

<span data-ttu-id="38d03-p102">La propriété **EditMode** est utile lorsqu'un processus de modification est interrompu, par une erreur lors de la validation, par exemple. La valeur de la propriété **EditMode** permet de déterminer la meilleure méthode à utiliser entre **[Update](recordset2-update-method-dao.md)** et **[CancelUpdate](recordset2-cancelupdate-method-dao.md)**.</span><span class="sxs-lookup"><span data-stu-id="38d03-p102">The **EditMode** property is useful when an editing process is interrupted, for example, by an error during validation. You can use the value of the **EditMode** property to determine whether you should use the **[Update](recordset2-update-method-dao.md)** or **[CancelUpdate](recordset2-cancelupdate-method-dao.md)** method.</span></span>

<span data-ttu-id="38d03-113">Vous pouvez également vérifier si le paramètre de propriété **[LockEdits](recordset2-lockedits-property-dao.md)** est défini sur **True** et si le paramètre de propriété **EditMode** est **dbEditInProgress** pour déterminer si la page active est verrouillée.</span><span class="sxs-lookup"><span data-stu-id="38d03-113">You can also check to see if the **[LockEdits](recordset2-lockedits-property-dao.md)** property setting is **True** and the **EditMode** property setting is **dbEditInProgress** to determine whether the current page is locked.</span></span>

## <a name="example"></a><span data-ttu-id="38d03-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="38d03-114">Example</span></span>

<span data-ttu-id="38d03-p103">Cet exemple illustre la valeur de la propriété **EditMode** sous différentes conditions. La fonction EditModeOutput est indispensable pour l'exécution de cette procédure.</span><span class="sxs-lookup"><span data-stu-id="38d03-p103">This example shows the value of the **EditMode** property under various conditions. The EditModeOutput function is required for this procedure to run.</span></span>

```vb
    Sub EditModeX() 
     
     Dim dbsNorthwind As Database 
     Dim rstEmployees As Recordset2 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     Set rstEmployees = _ 
     dbsNorthwind.OpenRecordset("Employees", _ 
     dbOpenDynaset) 
     
     ' Show the EditMode property under different editing 
     ' states. 
     With rstEmployees 
     EditModeOutput "Before any Edit or AddNew:", .EditMode 
     .Edit 
     EditModeOutput "After Edit:", .EditMode 
     .Update 
     EditModeOutput "After Update:", .EditMode 
     .AddNew 
     EditModeOutput "After AddNew:", .EditMode 
     .CancelUpdate 
     EditModeOutput "After CancelUpdate:", .EditMode 
     .Close 
     End With 
     
     dbsNorthwind.Close 
     
    End Sub 
     
    Function EditModeOutput(strTemp As String, _ 
     intEditMode As Integer) 
     
     ' Print report based on the value of the EditMode 
     ' property. 
     Debug.Print strTemp 
     Debug.Print " EditMode = "; 
     
     Select Case intEditMode 
     Case dbEditNone 
     Debug.Print "dbEditNone" 
     Case dbEditInProgress 
     Debug.Print "dbEditInProgress" 
     Case dbEditAdd 
     Debug.Print "dbEditAdd" 
     End Select 
     
    End Function
```