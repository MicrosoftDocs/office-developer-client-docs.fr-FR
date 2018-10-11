---
title: Field.ValidateOnSet Property (DAO)
TOCTitle: ValidateOnSet Property
ms:assetid: 00245a8a-a78f-b0a8-3eb3-11dd27873984
ms:mtpsurl: https://msdn.microsoft.com/library/Ff844720(v=office.15)
ms:contentKeyID: 48542898
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052929
f1_categories:
- Office.Version=v15
ms.openlocfilehash: e4b1297fe7c41efd49a457f3382f72c480f85363
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471343"
---
# <a name="fieldvalidateonset-property-dao"></a><span data-ttu-id="d8675-102">Field.ValidateOnSet Property (DAO)</span><span class="sxs-lookup"><span data-stu-id="d8675-102">Field.ValidateOnSet Property (DAO)</span></span>


<span data-ttu-id="d8675-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="d8675-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="d8675-104">Définit ou renvoie une valeur qui spécifie si la valeur d'un objet **[Field](field-object-dao.md)** est immédiatement validée quand la propriété **[Value](field-value-property-dao.md)** de l'objet est définie (espaces de travail Microsoft Access uniquement).</span><span class="sxs-lookup"><span data-stu-id="d8675-104">Sets or returns a value that specifies whether or not the value of a **[Field](field-object-dao.md)** object is immediately validated when the object's **[Value](field-value-property-dao.md)** property is set (Microsoft Access workspaces only).</span></span>

## <a name="syntax"></a><span data-ttu-id="d8675-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="d8675-105">Syntax</span></span>

<span data-ttu-id="d8675-106">*expression* . ValidateOnSet</span><span class="sxs-lookup"><span data-stu-id="d8675-106">*expression* .ValidateOnSet</span></span>

<span data-ttu-id="d8675-107">*expression* Variable qui représente un objet **Field** .</span><span class="sxs-lookup"><span data-stu-id="d8675-107">*expression* A variable that represents a **Field** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="d8675-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="d8675-108">Remarks</span></span>

<span data-ttu-id="d8675-109">Seuls les objets **Field** des objets **[Recordset](recordset-object-dao.md)** prennent en charge la propriété **ValidateOnSet** en lecture/écriture.</span><span class="sxs-lookup"><span data-stu-id="d8675-109">Only **Field** objects in **[Recordset](recordset-object-dao.md)** objects support the **ValidateOnSet** property as read/write.</span></span>

<span data-ttu-id="d8675-p101">Il peut être utile d'attribuer la valeur **True** à la propriété **ValidateOnSet** lorsqu'un utilisateur entre des enregistrements contenant un grand nombre de données Mémo. Si vous attendez que l'appel **[Update](recordset-update-method-dao.md)** valide les données, l'écriture des données Mémo dans la base de données risque de vous faire perdre un temps considérable inutilement s'il s'avère que ces données ne sont pas valides parce qu'elles enfreignent une règle de validation d'un autre champ.</span><span class="sxs-lookup"><span data-stu-id="d8675-p101">Setting the **ValidateOnSet** property to **True** can be useful in a situation when a user is entering records that include substantial Memo data. Waiting until the **[Update](recordset-update-method-dao.md)** call to validate the data can result in unnecessary time spent writing the lengthy Memo data to the database if it turns out that the data was invalid anyway because a validation rule was broken in another field.</span></span>

## <a name="example"></a><span data-ttu-id="d8675-112">Exemple</span><span class="sxs-lookup"><span data-stu-id="d8675-112">Example</span></span>

<span data-ttu-id="d8675-p102">Cet exemple utilise la propriété **ValidateOnSet** pour montrer comment intercepter des erreurs pendant l'entrée des données. La fonction ValidateData est obligatoire pour l'exécution de cette procédure.</span><span class="sxs-lookup"><span data-stu-id="d8675-p102">This example uses the **ValidateOnSet** property to demonstrate how one might trap for errors during data entry. The ValidateData function is required for this procedure to run.</span></span>

```vb
    Sub ValidateOnSetX() 
     
     Dim dbsNorthwind As Database 
     Dim fldDays As Field 
     Dim rstEmployees As Recordset 
     
     Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
     
     ' Create and append a new Field object to the Fields 
     ' collection of the Employees table. 
     Set fldDays = _ 
     dbsNorthwind.TableDefs!Employees.CreateField( _ 
     "DaysOfVacation", dbInteger, 2) 
     fldDays.ValidationRule = "BETWEEN 1 AND 20" 
     fldDays.ValidationText = _ 
     "Number must be between 1 and 20!" 
     dbsNorthwind.TableDefs!Employees.Fields.Append fldDays 
     
     Set rstEmployees = _ 
     dbsNorthwind.OpenRecordset("Employees") 
     
     With rstEmployees 
     
     Do While True 
     ' Add new record. 
     .AddNew 
     
     ' Get user input for three fields. Verify that the 
     ' data do not violate the validation rules for any 
     ' of the fields. 
     If ValidateData(!FirstName, _ 
     "Enter first name.") = False Then Exit Do 
     If ValidateData(!LastName, _ 
     "Enter last name.") = False Then Exit Do 
     If ValidateData(!DaysOfVacation, _ 
     "Enter days of vacation.") = False Then Exit Do 
     
     .Update 
     .Bookmark = .LastModified 
     Debug.Print !FirstName & " " & !LastName & _ 
     " - " & "DaysOfVacation = " & !DaysOfVacation 
     
     ' Delete new record because this is a demonstration. 
     .Delete 
     Exit Do 
     Loop 
     
     ' Cancel AddNew method if any of the validation rules 
     ' were broken. 
     If .EditMode <> dbEditNone Then .CancelUpdate 
     .Close 
     End With 
     
     ' Delete new field because this is a demonstration. 
     dbsNorthwind.TableDefs!Employees.Fields.Delete _ 
     fldDays.Name 
     dbsNorthwind.Close 
     
    End Sub 
     
    Function ValidateData(fldTemp As Field, _ 
     strMessage As String) As Boolean 
     
     Dim strInput As String 
     Dim errLoop As Error 
     
     ValidateData = True 
     ' ValidateOnSet is only read/write for Field objects in 
     ' Recordset objects. 
     fldTemp.ValidateOnSet = True 
     
     Do While True 
     strInput = InputBox(strMessage) 
     If strInput = "" Then Exit Do 
     ' Trap for errors when setting the Field value. 
     On Error GoTo Err_Data 
     If fldTemp.Type = dbInteger Then 
     fldTemp = Val(strInput) 
     Else 
     fldTemp = strInput 
     End If 
     On Error GoTo 0 
     If Not IsNull(fldTemp) Then Exit Do 
     Loop 
     
     If strInput = "" Then ValidateData = False 
     
     Exit Function 
     
    Err_Data: 
     
     If DBEngine.Errors.Count > 0 Then 
     ' Enumerate the Errors collection. The description 
     ' property of the last Error object will be set to 
     ' the ValidationText property of the relevant 
     ' field. 
     For Each errLoop In DBEngine.Errors 
     MsgBox "Error number: " & errLoop.Number & _ 
     vbCr & errLoop.Description 
     Next errLoop 
     End If 
     
     Resume Next 
     
    End Function
```