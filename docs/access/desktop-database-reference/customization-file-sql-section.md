---
title: Section SQL du fichier de personnalisation
TOCTitle: Customization File SQL Section
ms:assetid: 002c544f-fe1b-6aeb-ba9a-97b1e1159516
ms:mtpsurl: https://msdn.microsoft.com/library/JJ248776(v=office.15)
ms:contentKeyID: 48542901
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 09be671ba15727e3612ade78c5b54af12b1d1c57
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469431"
---
# <a name="customization-file-sql-section"></a><span data-ttu-id="adf99-102">Section SQL du fichier de personnalisation</span><span class="sxs-lookup"><span data-stu-id="adf99-102">Customization File SQL Section</span></span>


<span data-ttu-id="adf99-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="adf99-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="adf99-p101">La section **sql** peut contenir une nouvelle chaîne SQL qui remplace la chaîne de commande cliente. Si aucune chaîne SQL ne figure dans la section, celle-ci est ignorée.</span><span class="sxs-lookup"><span data-stu-id="adf99-p101">The **sql** section can contain a new SQL string that replaces the client command string. If there is no SQL string in the section, the section will be ignored.</span></span>

<span data-ttu-id="adf99-p102">La nouvelle chaîne SQL peut être *paramétrée*. En d'autres termes, les paramètres dans la chaîne SQL de la section **sql** (signalée par le caractère '?') peuvent être remplacés par les arguments correspondants dans un *identificateur* de la chaîne de commande cliente (il s'agit d'une liste séparée par des virgules entre parenthèses). L'identificateur et la liste des arguments ont un comportement similaire à celui d'un appel de fonction.</span><span class="sxs-lookup"><span data-stu-id="adf99-p102">The new SQL string may be *parameterized*. That is, parameters in the **sql** section SQL string (designated by the '?' character) can be replaced by corresponding arguments in an *identifier* in the client command string (designated by a comma-delimited list in parentheses). The identifier and argument list behave like a function call.</span></span>

<span data-ttu-id="adf99-109">Par exemple, supposons que la chaîne de commande cliente est « CustomerById (4) », l’en-tête de section SQL \[SQL CustomerByID\] , et la nouvelle chaîne de section SQL est « sélectionnez \* FROM Customers WHERE CustomerID = ? ».</span><span class="sxs-lookup"><span data-stu-id="adf99-109">For example, assume the client command string is "CustomerByID(4)" , the SQL section header is \[SQL CustomerByID\] , and the new SQL section string is "SELECT \* FROM Customers WHERE CustomerID = ?".</span></span> <span data-ttu-id="adf99-110">Le gestionnaire générera, l’en-tête de section SQL \[SQL CustomerByID\] , et la nouvelle chaîne de section SQL est « sélectionnez \* FROM Customers WHERE CustomerID = ? ».</span><span class="sxs-lookup"><span data-stu-id="adf99-110">The Handler will generate , the SQL section header is \[SQL CustomerByID\] , and the new SQL section string is "SELECT \* FROM Customers WHERE CustomerID = ?".</span></span> <span data-ttu-id="adf99-111">Le gestionnaire générera « sélectionnez \* FROM Customers WHERE CustomerID = 4 » et utilisera cette chaîne pour interroger la source de données.</span><span class="sxs-lookup"><span data-stu-id="adf99-111">The Handler will generate "SELECT \* FROM Customers WHERE CustomerID = 4" and use that string to query the data source.</span></span>

<span data-ttu-id="adf99-112">Si la nouvelle instruction SQL est une chaîne vide (""), la section est ignorée.</span><span class="sxs-lookup"><span data-stu-id="adf99-112">If the new SQL statement is the null string (""), then the section is ignored.</span></span>

<span data-ttu-id="adf99-p104">Si la chaîne de la nouvelle instruction SQL n'est pas valide, l'exécution de l'instruction échoue. Le paramètre client est ignoré. Vous pouvez avoir recours à cette programmation intentionnellement pour « désactiver » toutes les commandes SQL clientes en spécifiant :</span><span class="sxs-lookup"><span data-stu-id="adf99-p104">If the new SQL statement string is not valid, then the execution of the statement will fail. The client parameter is effectively ignored. You can do this intentionally to "turn off" all client SQL commands by specifying:</span></span>

```sql 
 
[SQL default] 
SQL = " " 
```

## <a name="syntax"></a><span data-ttu-id="adf99-116">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="adf99-116">Syntax</span></span>

<span data-ttu-id="adf99-117">Une chaîne d'entrée SQL de remplacement a la forme suivante :</span><span class="sxs-lookup"><span data-stu-id="adf99-117">A replacement SQL string entry is of the form:</span></span>

<span data-ttu-id="adf99-118">**SQL = \* chaîne-SQL**\*</span><span class="sxs-lookup"><span data-stu-id="adf99-118">**SQL=\*sqlString**\*</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="adf99-119">Élément</span><span class="sxs-lookup"><span data-stu-id="adf99-119">Part</span></span></p></th>
<th><p><span data-ttu-id="adf99-120">Description</span><span class="sxs-lookup"><span data-stu-id="adf99-120">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="adf99-121"><strong>SQL</strong></span><span class="sxs-lookup"><span data-stu-id="adf99-121"><strong>SQL</strong></span></span></p></td>
<td><p><span data-ttu-id="adf99-122">Chaîne littérale qui indique qu'il s'agit d'une entrée de section SQL.</span><span class="sxs-lookup"><span data-stu-id="adf99-122">A literal string that indicates this is an SQL section entry.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="adf99-123"><strong><em>chaîne-SQL</em></strong></span><span class="sxs-lookup"><span data-stu-id="adf99-123"><strong><em>sqlString</em></strong></span></span></p></td>
<td><p><span data-ttu-id="adf99-124">Chaîne SQL qui remplace la chaîne cliente.</span><span class="sxs-lookup"><span data-stu-id="adf99-124">An SQL string that replaces the client string.</span></span></p></td>
</tr>
</tbody>
</table>
