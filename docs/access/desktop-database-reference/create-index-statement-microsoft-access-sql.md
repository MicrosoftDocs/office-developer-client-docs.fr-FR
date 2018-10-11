---
title: CREATE INDEX, instruction (Microsoft Access SQL)
TOCTitle: CREATE INDEX Statement (Microsoft Access SQL)
ms:assetid: c5919ef4-a08d-df06-7078-5331adbcb45c
ms:mtpsurl: https://msdn.microsoft.com/library/Ff823109(v=office.15)
ms:contentKeyID: 48547612
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- jetsql40.chm5277562
f1_categories:
- Office.Version=v15
ms.openlocfilehash: ab501348d19ad8577bf1a55a3f37c6c3923381b1
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469474"
---
# <a name="create-index-statement-microsoft-access-sql"></a><span data-ttu-id="51db3-102">CREATE INDEX, instruction (Microsoft Access SQL)</span><span class="sxs-lookup"><span data-stu-id="51db3-102">CREATE INDEX Statement (Microsoft Access SQL)</span></span>

<span data-ttu-id="51db3-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="51db3-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="51db3-104">Crée un nouvel index sur une table existante.</span><span class="sxs-lookup"><span data-stu-id="51db3-104">Creates a new index on an existing table.</span></span>

> [!NOTE]
> <span data-ttu-id="51db3-p101">[!REMARQUE] Pour les bases de données autres que celles de type Microsoft Access, le moteur de base de données Microsoft Access ne prend pas en charge CREATE INDEX (sauf pour créer un pseudo index sur une table liée ODBC), ni les instructions du langage de définition de données (DDL). Pour cela, utilisez les méthodes Create DAO . Pour plus d'informations, voir la section Notes.</span><span class="sxs-lookup"><span data-stu-id="51db3-p101">For non-Microsoft Access atabse engine databases, the Microsoft Access database engine does not support the use of CREATE INDEX (except to create a pseudo index on an ODBC linked table) or any of the data definition language (DDL) statements. Use the DAO Create methods instead. For more information see the Remarks section.</span></span>

## <a name="syntax"></a><span data-ttu-id="51db3-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="51db3-108">Syntax</span></span>

<span data-ttu-id="51db3-109">CRÉER \[ UNIQUE \] INDEX *index* ON *table* (*champ* \[ASC | DESC\]\[, *champ* \[ASC | DESC\],... \]) \[WITH {PRIMARY | DISALLOW NULL | IGNORE NULL}\]</span><span class="sxs-lookup"><span data-stu-id="51db3-109">CREATE \[ UNIQUE \] INDEX *index* ON *table* (*field* \[ASC|DESC\]\[, *field* \[ASC|DESC\], …\]) \[WITH { PRIMARY | DISALLOW NULL | IGNORE NULL }\]</span></span>

<span data-ttu-id="51db3-110">L'instruction CREATE INDEX est composée des arguments suivants :</span><span class="sxs-lookup"><span data-stu-id="51db3-110">The CREATE INDEX statement has these parts:</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="51db3-111">Argument</span><span class="sxs-lookup"><span data-stu-id="51db3-111">Part</span></span></p></th>
<th><p><span data-ttu-id="51db3-112">Description</span><span class="sxs-lookup"><span data-stu-id="51db3-112">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="51db3-113"><em>index</em></span><span class="sxs-lookup"><span data-stu-id="51db3-113"><em>index</em></span></span></p></td>
<td><p><span data-ttu-id="51db3-114">Nom de l'index à créer.</span><span class="sxs-lookup"><span data-stu-id="51db3-114">The name of the index to be created.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="51db3-115"><em>table</em></span><span class="sxs-lookup"><span data-stu-id="51db3-115"><em>table</em></span></span></p></td>
<td><p><span data-ttu-id="51db3-116">Nom de la table existante qui contiendra l'index.</span><span class="sxs-lookup"><span data-stu-id="51db3-116">The name of the existing table that will contain the index.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="51db3-117"><em>champ</em></span><span class="sxs-lookup"><span data-stu-id="51db3-117"><em>field</em></span></span></p></td>
<td><p><span data-ttu-id="51db3-p102">Nom du champ ou des champs à indexer. Pour créer un index d'un seul champ, indiquez le nom du champ entre parenthèses suivi du nom de la table. Pour créer un index de plusieurs champs, indiquez le nom de chaque champ à inclure dans l'index. Pour créer des index décroissants, utilisez le mot réservé DESC, sinon les index sont considérés comme croissants.</span><span class="sxs-lookup"><span data-stu-id="51db3-p102">The name of the field or fields to be indexed. To create a single-field index, list the field name in parentheses following the table name. To create a multiple-field index, list the name of each field to be included in the index. To create descending indexes, use the DESC reserved word; otherwise, indexes are assumed to be ascending.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="51db3-122">Notes</span><span class="sxs-lookup"><span data-stu-id="51db3-122">Remarks</span></span>

<span data-ttu-id="51db3-123">Pour interdire les valeurs dupliquées dans les champs indexés de différent enregistrements, utilisez le mot réservé UNIQUE.</span><span class="sxs-lookup"><span data-stu-id="51db3-123">To prohibit duplicate values in the indexed field or fields of different records, use the UNIQUE reserved word.</span></span>

<span data-ttu-id="51db3-p103">Vous pouvez appliquer des règles de validation dans la clause WITH facultative. Vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="51db3-p103">In the optional WITH clause you can enforce data validation rules. You can:</span></span>

- <span data-ttu-id="51db3-126">interdire les entrées Null dans les champs indexés de nouveaux enregistrements à l'aide de l'option DISALLOW NULL ;</span><span class="sxs-lookup"><span data-stu-id="51db3-126">Prohibit Null entries in the indexed field or fields of new records by using the DISALLOW NULL option.</span></span>

- <span data-ttu-id="51db3-127">empêcher que des enregistrements avec des valeurs **Null** dans les champs indexés soient inclus dans l'index à l'aide de l'option IGNORE NULL ;</span><span class="sxs-lookup"><span data-stu-id="51db3-127">Prevent records with **Null** values in the indexed field or fields from being included in the index by using the IGNORE NULL option.</span></span>

- <span data-ttu-id="51db3-p104">désigner les champs indexés en tant que clé primaire en utilisant le mot réservé PRIMARY. Ceci implique que la clé soit unique, aussi vous pouvez omettre le mot réservé UNIQUE.</span><span class="sxs-lookup"><span data-stu-id="51db3-p104">Designate the indexed field or fields as the primary key by using the PRIMARY reserved word. This implies that the key is unique, so you can omit the UNIQUE reserved word.</span></span>

<span data-ttu-id="51db3-p105">Vous pouvez utiliser CREATE INDEX pour créer un pseudo index sur une table liée dans une source de données ODBC, telle que Microsoft® SQL Server qui ne comporte pas d'index. Vous n'avez pas besoin d'une autorisation d'accès particulière au serveur distant pour créer un pseudo index ; la base de données distante ignore l'existence du pseudo index qui n'a par conséquent aucune incidence sur celle-ci. Utilisez la même syntaxe pour les tables liées et natives. Créer un pseudo index sur une table qui est normalement en lecture seule peut s'avérer particulièrement utile.</span><span class="sxs-lookup"><span data-stu-id="51db3-p105">You can use CREATE INDEX to create a pseudo index on a linked table in an ODBC data source, such as Microsoft® SQL Server™, that does not already have an index. You do not need permission or access to the remote server to create a pseudo index, and the remote database is unaware of and unaffected by the pseudo index. You use the same syntax for both linked and native tables. Creating a pseudo-index on a table that would ordinarily be read-only can be especially useful.</span></span>

<span data-ttu-id="51db3-134">Vous pouvez également utiliser l'instruction [ALTER TABLE](alter-table-statement-microsoft-access-sql.md) pour ajouter un index d'un seul champ ou de plusieurs champs à une table et l'instruction ALTER TABLE ou [DROP](drop-statement-microsoft-access-sql.md) pour supprimer un index créé avec ALTER TABLE ou CREATE INDEX.</span><span class="sxs-lookup"><span data-stu-id="51db3-134">You can also use the [ALTER TABLE](alter-table-statement-microsoft-access-sql.md) statement to add a single- or multiple-field index to a table, and you can use the ALTER TABLE statement or the [DROP](drop-statement-microsoft-access-sql.md) statement to remove an index created with ALTER TABLE or CREATE INDEX.</span></span>

> [!NOTE]
> <span data-ttu-id="51db3-135">[!REMARQUE] N'utilisez pas le mot réservé PRIMARY lorsque vous créez un nouvel index sur une table qui comporte déjà une clé primaire ; si vous le faites, une erreur se produit.</span><span class="sxs-lookup"><span data-stu-id="51db3-135">Do not use the PRIMARY reserved word when you create a new index on a table that already has a primary key; if you do, an error occurs.</span></span>

## <a name="example"></a><span data-ttu-id="51db3-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="51db3-136">Example</span></span>

<span data-ttu-id="51db3-137">Dans cet exemple, un index comportant les champs Home Phone et Extension dans la table Employees est créé.</span><span class="sxs-lookup"><span data-stu-id="51db3-137">This example creates an index consisting of the fields Home Phone and Extension in the Employees table.</span></span>

```vb
    Sub CreateIndexX1() 
     
        Dim dbs As Database 
     
        ' Modify this line to include the path to Northwind 
        ' on your computer. 
        Set dbs = OpenDatabase("Northwind.mdb") 
     
        ' Create the NewIndex index on the Employees table. 
        dbs.Execute "CREATE INDEX NewIndex ON Employees " _ 
            & "(HomePhone, Extension);" 
     
        dbs.Close 
     
    End Sub 
```

<br/>

<span data-ttu-id="51db3-p106">Dans cet exemple, un index est créé sur la table Customers utilisant le champ CustomerID. Deux enregistrements ne peuvent pas avoir les mêmes données dans le champ CustomerID et aucune valeur Null n'est autorisée.</span><span class="sxs-lookup"><span data-stu-id="51db3-p106">This example creates an index on the Customers table using the CustomerID field. No two records can have the same data in the CustomerID field, and no Null values are allowed.</span></span>

```vb
    Sub CreateIndexX2() 
     
        Dim dbs As Database 
     
        ' Modify this line to include the path to Northwind 
        ' on your computer. 
        Set dbs = OpenDatabase("Northwind.mdb") 
     
        ' Create a unique index, CustID, on the  
        ' CustomerID field. 
        dbs.Execute "CREATE UNIQUE INDEX CustID " _ 
            & "ON Customers (CustomerID) " _ 
            & "WITH DISALLOW NULL;" 
     
        dbs.Close 
     
    End Sub
```