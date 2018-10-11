---
title: TRANSFORM, instruction (Microsoft Access SQL)
TOCTitle: TRANSFORM Statement (Microsoft Access SQL)
ms:assetid: 419770b1-c833-959d-a84d-56c68764799f
ms:mtpsurl: https://msdn.microsoft.com/library/Ff192901(v=office.15)
ms:contentKeyID: 48544455
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- jetsql40.chm5277581
f1_categories:
- Office.Version=v15
ms.openlocfilehash: 6d05f278e38cc8cf132cf06605703dfa99eb8728
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471516"
---
# <a name="transform-statement-microsoft-access-sql"></a><span data-ttu-id="f34d4-102">TRANSFORM, instruction (Microsoft Access SQL)</span><span class="sxs-lookup"><span data-stu-id="f34d4-102">TRANSFORM Statement (Microsoft Access SQL)</span></span>

<span data-ttu-id="f34d4-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="f34d4-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="f34d4-104">Crée une requête Analyse croisée.</span><span class="sxs-lookup"><span data-stu-id="f34d4-104">Creates a crosstab query.</span></span>

## <a name="syntax"></a><span data-ttu-id="f34d4-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="f34d4-105">Syntax</span></span>

<span data-ttu-id="f34d4-106">TRANSFORM *fonctionagrégationinstructionselect* PIVOT *pivotfield* \[IN (*valeur1*\[, *valeur2*\[,... \]\])\]</span><span class="sxs-lookup"><span data-stu-id="f34d4-106">TRANSFORM *aggfunctionselectstatement* PIVOT *pivotfield* \[IN (*value1*\[, *value2*\[, …\]\])\]</span></span>

<span data-ttu-id="f34d4-107">L'instruction TRANSFORM est composée des arguments suivants :</span><span class="sxs-lookup"><span data-stu-id="f34d4-107">The TRANSFORM statement has these parts:</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="f34d4-108">Argument</span><span class="sxs-lookup"><span data-stu-id="f34d4-108">Part</span></span></p></th>
<th><p><span data-ttu-id="f34d4-109">Description</span><span class="sxs-lookup"><span data-stu-id="f34d4-109">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="f34d4-110"><em>fonctionagrégation</em></span><span class="sxs-lookup"><span data-stu-id="f34d4-110"><em>aggfunction</em></span></span></p></td>
<td><p><span data-ttu-id="f34d4-111"><a href="sql-aggregate-functions-sql.md">Fonction d’agrégation SQL</a> agissant sur les données sélectionnées.</span><span class="sxs-lookup"><span data-stu-id="f34d4-111">An <a href="sql-aggregate-functions-sql.md">SQL aggregate function</a> that operates on the selected data.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="f34d4-112"><em>instructionselect</em></span><span class="sxs-lookup"><span data-stu-id="f34d4-112"><em>selectstatement</em></span></span></p></td>
<td><p><span data-ttu-id="f34d4-113">Une instruction <a href="select-statement-microsoft-access-sql.md">SELECT</a>.</span><span class="sxs-lookup"><span data-stu-id="f34d4-113">A <a href="select-statement-microsoft-access-sql.md">SELECT</a> statement.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="f34d4-114"><em>champpivot</em></span><span class="sxs-lookup"><span data-stu-id="f34d4-114"><em>pivotfield</em></span></span></p></td>
<td><p><span data-ttu-id="f34d4-115">Champ ou expression que vous voulez utilisez pour créer les en-têtes de colonne dans le jeu de résultats de la requête.</span><span class="sxs-lookup"><span data-stu-id="f34d4-115">The field or expression you want to use to create column headings in the query's result set.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="f34d4-116"><em>valeur1</em>, <em>valeur2</em></span><span class="sxs-lookup"><span data-stu-id="f34d4-116"><em>value1</em>, <em>value2</em></span></span></p></td>
<td><p><span data-ttu-id="f34d4-117">Valeurs fixes utilisées pour créer des en-têtes de colonne.</span><span class="sxs-lookup"><span data-stu-id="f34d4-117">Fixed values used to create column headings.</span></span></p></td>
</tr>
</tbody>
</table>

## <a name="remarks"></a><span data-ttu-id="f34d4-118">Notes</span><span class="sxs-lookup"><span data-stu-id="f34d4-118">Remarks</span></span>

<span data-ttu-id="f34d4-119">Lorsque vous synthétisez des données à l'aide d'une requête Analyse croisée, vous sélectionnez des valeurs à partir de champs ou d'expressions spécifiés comme en-têtes de colonne. Ainsi, vous pouvez voir les données sous un format plus condensé qu'avec une requête Sélection.</span><span class="sxs-lookup"><span data-stu-id="f34d4-119">When you summarize data using a crosstab query, you select values from specified fields or expressions as column headings so you can view data in a more compact format than with a select query.</span></span>

<span data-ttu-id="f34d4-p101">L'instruction TRANSFORM est facultative mais, si elle est employée, elle doit venir en première position dans une chaîne SQL. Elle précède une instruction SELECT, qui spécifie les champs servant d'en-têtes de ligne ainsi qu'une clause [GROUP BY](https://msdn.microsoft.com/library/ff837271\(v=office.15\)) spécifiant le mode de regroupement des lignes. Vous pouvez éventuellement ajouter d'autres clauses, par exemple une clause [WHERE](https://msdn.microsoft.com/library/ff195245\(v=office.15\)) précisant des critères de sélection ou de tri supplémentaires. Vous pouvez aussi utiliser des sous-requêtes comme prédicats, particulièrement ceux dans la clause WHERE, dans une requête Analyse croisée.</span><span class="sxs-lookup"><span data-stu-id="f34d4-p101">TRANSFORM is optional but when included is the first statement in an SQL string. It precedes a SELECT statement that specifies the fields used as row headings and a [GROUP BY](https://msdn.microsoft.com/library/ff837271\(v=office.15\)) clause that specifies row grouping. Optionally, you can include other clauses, such as [WHERE](https://msdn.microsoft.com/library/ff195245\(v=office.15\)), that specify additional selection or sorting criteria. You can also use subqueries as predicates — specifically, those in the WHERE clause — in a crosstab query.</span></span>

<span data-ttu-id="f34d4-p102">Les valeurs renvoyées dans *pivotfield* sont utilisées comme en-têtes de colonne dans le jeu de résultats de la requête. Par exemple, créer une requête Analyse croisée en prenant le mois de vente comme base de sélection pour obtenir les chiffres des ventes produira douze colonnes. Vous pouvez restreindre *pivotfield* pour que les en-têtes ne soient créés qu'à partir des valeurs fixes (*valeur1*, *valeur2* ) répertoriées dans la clause IN facultative. Vous pouvez également ajouter des valeurs fixes sans données correspondantes de manière à créer des colonnes supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="f34d4-p102">The values returned in *pivotfield* are used as column headings in the query's result set. For example, pivoting the sales figures on the month of the sale in a crosstab query would create 12 columns. You can restrict *pivotfield* to create headings from fixed values (*value1*, *value2* ) listed in the optional IN clause. You can also include fixed values for which no data exists to create additional columns.</span></span>

## <a name="example"></a><span data-ttu-id="f34d4-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="f34d4-128">Example</span></span>

<span data-ttu-id="f34d4-p103">Dans cet exemple la clause SQL TRANSFORM est utilisée pour créer une requête Analyse croisée qui indique le nombre de commandes prises par chaque employé pour chaque trimestre de 1994. La fonction SQLTRANSFORMOutput est nécessaire pour que la procédure puisse s'exécuter.</span><span class="sxs-lookup"><span data-stu-id="f34d4-p103">This example uses the SQL TRANSFORM clause to create a crosstab query showing the number of orders taken by each employee for each calendar quarter of 1994. The SQLTRANSFORMOutput function is required for this procedure to run.</span></span>

```vb
    Sub TransformX1() 
     
        Dim dbs As Database 
        Dim strSQL As String 
        Dim qdfTRANSFORM As QueryDef 
     
        strSQL = "PARAMETERS prmYear SHORT; TRANSFORM " _ 
            & "Count(OrderID) " _ 
            & "SELECT FirstName & "" "" & LastName AS " _ 
            & "FullName FROM Employees INNER JOIN Orders " _ 
            & "ON Employees.EmployeeID = " _ 
            & "Orders.EmployeeID WHERE DatePart " _ 
            & "(""yyyy"", OrderDate) = [prmYear] " 
       
           strSQL = strSQL & "GROUP BY FirstName & " _ 
            & """ "" & LastName " _ 
            & "ORDER BY FirstName & "" "" & LastName " _ 
            & "PIVOT DatePart(""q"", OrderDate)" 
         
        ' Modify this line to include the path to Northwind 
        ' on your computer. 
        Set dbs = OpenDatabase("Northwind.mdb") 
     
        Set qdfTRANSFORM = dbs.CreateQueryDef _ 
            ("", strSQL) 
         
        SQLTRANSFORMOutput qdfTRANSFORM, 1994 
         
        dbs.Close 
     
    End Sub
```

<br/>

<span data-ttu-id="f34d4-p104">Dans cet exemple, la clause SQL TRANSFORM est utilisée pour créer une requête Analyse croisée légèrement plus complexe qui indique le montant total en dollar des commandes prises par chaque employé pour chaque trimestre de 1994. La fonction SQLTRANSFORMOutput est nécessaire pour que la procédure puisse s'exécuter..</span><span class="sxs-lookup"><span data-stu-id="f34d4-p104">This example uses the SQL TRANSFORM clause to create a slightly more complex crosstab query showing the total dollar amount of orders taken by each employee for each calendar quarter of 1994. The SQLTRANSFORMOutput function is required for this procedure to run.</span></span>

```vb
    Sub TransformX2() 
     
        Dim dbs As Database 
        Dim strSQL As String 
        Dim qdfTRANSFORM As QueryDef 
     
        strSQL = "PARAMETERS prmYear SMALLINT; TRANSFORM " _ 
            & "Sum(Subtotal) SELECT FirstName & "" """ _ 
            & "& LastName AS FullName " _ 
            & "FROM Employees INNER JOIN " _ 
            & "(Orders INNER JOIN [Order Subtotals] " _ 
            & "ON Orders.OrderID = " _ 
            & "[Order Subtotals].OrderID) " _ 
            & "ON Employees.EmployeeID = " _ 
            & "Orders.EmployeeID WHERE DatePart" _ 
            & "(""yyyy"", OrderDate) = [prmYear] " 
        
           strSQL = strSQL & "GROUP BY FirstName & "" """ _ 
            & "& LastName " _ 
            & "ORDER BY FirstName & "" "" & LastName " _ 
            & "PIVOT DatePart(""q"",OrderDate)"         
             
        ' Modify this line to include the path to Northwind 
        ' on your computer. 
        Set dbs = OpenDatabase("Northwind.mdb") 
     
        Set qdfTRANSFORM = dbs.CreateQueryDef _ 
            ("", strSQL) 
         
        SQLTRANSFORMOutput qdfTRANSFORM, 1994 
         
        dbs.Close 
     
    End Sub 
     
    Function SQLTRANSFORMOutput(qdfTemp As QueryDef, _ 
        intYear As Integer) 
         
        Dim rstTRANSFORM As Recordset 
        Dim fldLoop As Field 
        Dim booFirst As Boolean 
     
        qdfTemp.PARAMETERS!prmYear = intYear 
        Set rstTRANSFORM = qdfTemp.OpenRecordset() 
         
        Debug.Print qdfTemp.SQL 
        Debug.Print 
        Debug.Print , , "Quarter" 
     
        With rstTRANSFORM 
            booFirst = True 
            For Each fldLoop In .Fields 
                If booFirst = True Then 
                    Debug.Print fldLoop.Name 
                    Debug.Print , ; 
                    booFirst = False 
                Else 
                    Debug.Print , fldLoop.Name; 
                End If 
            Next fldLoop 
            Debug.Print 
             
            Do While Not .EOF 
                booFirst = True 
                For Each fldLoop In .Fields 
                    If booFirst = True Then 
                        Debug.Print fldLoop 
                        Debug.Print , ; 
                        booFirst = False 
                    Else 
                        Debug.Print , fldLoop; 
                    End If 
                Next fldLoop 
                Debug.Print 
                .MoveNext 
            Loop 
        End With 
         
    End Function
```