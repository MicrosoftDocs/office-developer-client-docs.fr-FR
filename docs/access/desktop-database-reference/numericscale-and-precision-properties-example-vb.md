---
title: NumericScale et Precision, propriétés - Exemple (VB)
TOCTitle: NumericScale and Precision Properties Example (VB)
ms:assetid: 728a76a3-1f80-935b-b6c7-94255ffe0160
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249462(v=office.15)
ms:contentKeyID: 48545610
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 5c1787ba817f04fa7f2f90a25be37cb3d89dbca3
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470053"
---
# <a name="numericscale-and-precision-properties-example-vb"></a><span data-ttu-id="306f0-102">NumericScale et Precision, propriétés - Exemple (VB)</span><span class="sxs-lookup"><span data-stu-id="306f0-102">NumericScale and Precision Properties Example (VB)</span></span>


<span data-ttu-id="306f0-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="306f0-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="306f0-104">Cet exemple illustre les propriétés [NumericScale](numericscale-property-adox.md) et [Precision](precision-property-adox.md) de l'objet [Column](column-object-adox.md).</span><span class="sxs-lookup"><span data-stu-id="306f0-104">This example demonstrates the [NumericScale](numericscale-property-adox.md) and [Precision](precision-property-adox.md) properties of the [Column](column-object-adox.md) object.</span></span> <span data-ttu-id="306f0-105">Ce code affiche leur valeur pour la table **Order Details** de la base de données *Northwind* .</span><span class="sxs-lookup"><span data-stu-id="306f0-105">This code displays their value for the **Order Details** table of the *Northwind* database.</span></span>

```vb 
 
' BeginNumericScalePrecVB 
Sub Main() 
 On Error GoTo NumericScalePrecXError 
 
 Dim cnn As New ADODB.Connection 
 Dim cat As New ADOX.Catalog 
 Dim tblOD As ADOX.Table 
 Dim colLoop As ADOX.Column 
 
 ' Connect the catalog. 
 cnn.Open "Provider='Microsoft.Jet.OLEDB.4.0';" & _ 
 "data source='c:\Program Files\" & _ 
 "Microsoft Office\Office\Samples\Northwind.mdb';" 
 Set cat.ActiveConnection = cnn 
 
 ' Retrieve the Order Details table 
 Set tblOD = cat.Tables("Order Details") 
 
 ' Display numeric scale and precision of 
 ' small integer fields. 
 For Each colLoop In tblOD.Columns 
 If colLoop.Type = adSmallInt Then 
 MsgBox "Column: " & colLoop.Name & vbCr & _ 
 "Numeric scale: " & _ 
 colLoop.NumericScale & vbCr & _ 
 "Precision: " & colLoop.Precision 
 End If 
 Next colLoop 
 
 'Clean up 
 cnn.Close 
 Set cat = Nothing 
 Set cnn = Nothing 
 Exit Sub 
 
NumericScalePrecXError: 
 Set cat = Nothing 
 
 If Not cnn Is Nothing Then 
 If cnn.State = adStateOpen Then cnn.Close 
 End If 
 Set cnn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source & "-->" & Err.Description, , "Error" 
 End If 
End Sub 
' EndNumericScalePrecVB 
```
