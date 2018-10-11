---
title: CreateRecordset, méthode - Exemple (VB)
TOCTitle: CreateRecordset Method Example (VB)
ms:assetid: c1c05c91-0c74-1f30-7ead-6b52f0b4906c
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249943(v=office.15)
ms:contentKeyID: 48547536
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: c99db7e319d6678b6db60cc1877a43a1728a00c2
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470502"
---
# <a name="createrecordset-method-example-vb"></a><span data-ttu-id="a61b3-102">CreateRecordset, méthode - Exemple (VB)</span><span class="sxs-lookup"><span data-stu-id="a61b3-102">CreateRecordset Method Example (VB)</span></span>


<span data-ttu-id="a61b3-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="a61b3-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="a61b3-p101">Vous pouvez créer un objet [Recordset](recordset-object-ado.md) et spécifier les informations de colonne, puis insérer des données dans cet objet **Recordset**; l'ensemble de lignes sous-jacent place les insertions dans la mémoire tampon.</span><span class="sxs-lookup"><span data-stu-id="a61b3-p101">You can create a [Recordset](recordset-object-ado.md) object and specify the column information. You can then insert data into the **Recordset** object; the underlying rowset buffers the inserts.</span></span>

<span data-ttu-id="a61b3-p102">L'exemple de code suivant montre comment définir un objet **Recordset** à l'aide de l'objet [RDSServer.DataFactory](datafactory-object-rdsserver.md). Vous pouvez également effectuer cette opération avec l'objet [RDS.DataControl](datacontrol-object-rds.md).</span><span class="sxs-lookup"><span data-stu-id="a61b3-p102">The following code example shows how to define a **Recordset** by using the [RDSServer.DataFactory](datafactory-object-rdsserver.md) object. You can also do this with the [RDS.DataControl](datacontrol-object-rds.md) object.</span></span>

```vb 
 
'BeginRsDefineShapeVB 
Sub Main() 
 On Error GoTo ErrorHandler 
 
 Dim ADF As RDSServer.DataFactory 
 Dim vntRecordShape(3) 
 Dim vntField1Shape(3) 
 Dim vntField2Shape(3) 
 Dim vntField3Shape(3) 
 Dim vntField4Shape(3) 
 
 Set ADF = New RDSServer.DataFactory 
 
 ' For each field, specify the name, 
 ' type, size, and nullability. 
 
 vntField1Shape(0) = "Name" ' Column name. 
 vntField1Shape(1) = CInt(129) ' Column type. 
 vntField1Shape(2) = CInt(40) ' Column size. 
 vntField1Shape(3) = False ' Nullable? 
 
 vntField2Shape(0) = "Age" 
 vntField2Shape(1) = CInt(3) 
 vntField2Shape(2) = CInt(-1) 
 vntField2Shape(3) = True 
 
 vntField3Shape(0) = "DateOfBirth" 
 vntField3Shape(1) = CInt(7) 
 vntField3Shape(2) = CInt(-1) 
 vntField3Shape(3) = True 
 
 vntField4Shape(0) = "Balance" 
 vntField4Shape(1) = CInt(6) 
 vntField4Shape(2) = CInt(-1) 
 vntField4Shape(3) = True 
 
 ' Put all fields into an array of arrays. 
 vntRecordShape(0) = vntField1Shape 
 vntRecordShape(1) = vntField2Shape 
 vntRecordShape(2) = vntField3Shape 
 vntRecordShape(3) = vntField4Shape 
 
 ' Use the RDSServer.DataFactory to create an empty 
 ' recordset. It takes an array of variants where 
 ' every element is itself another array of 
 ' variants, one for every column required in the 
 ' recordset. 
 ' The elements of the inner array are the column's 
 ' name, type, size, and nullability. 
 ' 
 ' NOTE: You could just use the RDS.DataControl object 
 ' instead of the RDSServer.DataFactory object. In 
 ' that case, the following code would be Set NewRS 
 ' = ADC1.CreateRecordset(vntRecordShape) 
 Dim NewRs As ADODB.Recordset 
 Set NewRs = ADF.CreateRecordSet(vntRecordShape) 
 
 Dim fields(3) 
 fields(0) = vntField1Shape(0) 
 fields(1) = vntField2Shape(0) 
 fields(2) = vntField3Shape(0) 
 fields(3) = vntField4Shape(0) 
 
 ' Populate the new recordset with data values. 
 Dim fieldVals(3) 
 
 ' Use AddNew to add the records. 
 fieldVals(0) = "Joe" 
 fieldVals(1) = 5 
 fieldVals(2) = CDate(#1/5/1996#) 
 fieldVals(3) = 123.456 
 NewRs.AddNew fields, fieldVals 
 
 fieldVals(0) = "Mary" 
 fieldVals(1) = 6 
 fieldVals(2) = CDate(#6/5/1996#) 
 fieldVals(3) = 31 
 NewRs.AddNew fields, fieldVals 
 
 fieldVals(0) = "Alex" 
 fieldVals(1) = 13 
 fieldVals(2) = CDate(#1/6/1996#) 
 fieldVals(3) = 34.0001 
 NewRs.AddNew fields, fieldVals 
 
 fieldVals(0) = "Susan" 
 fieldVals(1) = 13 
 fieldVals(2) = CDate(#8/6/1996#) 
 fieldVals(3) = 0# 
 NewRs.AddNew fields, fieldVals 
 
 NewRs.MoveFirst 
 
 ' Set the newly created and populated Recordset to 
 ' the SourceRecordset property of the 
 ' RDS.DataControl to bind to visual controls 
 Dim ADC1 As RDS.DataControl 
 Set ADC1 = New RDS.DataControl 
 Set ADC1.SourceRecordset = NewRs 
 
 'Clean up 
 If NewRs.State = adStateOpen Then NewRs.Close 
 Set NewRs = Nothing 
 Set ADC1 = Nothing 
 Set ADF = Nothing 
 Exit Sub 
 
ErrorHandler: 
 ' clean up 
 If Not NewRs Is Nothing Then 
 If NewRs.State = adStateOpen Then NewRs.Close 
 End If 
 Set NewRs = Nothing 
 Set ADC1 = Nothing 
 Set ADF = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
End Sub 
'EndRsDefineShapeVB 
```
