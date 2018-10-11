---
title: Création de jeux d'enregistrements hiérarchiques
TOCTitle: Fabricating Hierarchical Recordsets
ms:assetid: 0a6e41ba-015e-c07e-8876-1e744256b876
ms:mtpsurl: https://msdn.microsoft.com/library/JJ248836(v=office.15)
ms:contentKeyID: 48543153
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 81302ba1c18645a51913cb92179f4b61f3c32ce4
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471635"
---
# <a name="fabricating-hierarchical-recordsets"></a><span data-ttu-id="7088c-102">Création de jeux d'enregistrements hiérarchiques</span><span class="sxs-lookup"><span data-stu-id="7088c-102">Fabricating Hierarchical Recordsets</span></span>


<span data-ttu-id="7088c-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="7088c-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="7088c-104">L'exemple ci-après indique comment créer un jeu d'enregistrements hiérarchique sans source de données sous-jacente en utilisant la syntaxe de mise en forme des données pour définir les colonnes des objets **Recordset** parents, enfants ou petits-enfants.</span><span class="sxs-lookup"><span data-stu-id="7088c-104">The following example shows how to fabricate a hierarchical Recordset without an underlying data source by using the data shaping grammar to define columns for parent, child, and grandchild **Recordsets**.</span></span>

<span data-ttu-id="7088c-p101">Pour créer un **jeu d'enregistrements**, hiérarchique, vous devez spécifier le service de mise en forme des données Microsoft pour OLE DB (MSDataShape). Vous pouvez spécifier la valeur de fournisseur de données NONE dans le paramètre de chaîne de connexion de la méthode [Open](connection-object-ado.md) de l'objet [Connection](open-method-ado-connection.md) . Pour plus d'informations, reportez-vous à la section [Fournisseurs requis pour la mise en forme des données](required-providers-for-data-shaping.md).</span><span class="sxs-lookup"><span data-stu-id="7088c-p101">To fabricate a hierarchical **Recordset**, you must specify the Microsoft Data Shaping Service for OLE DB (MSDataShape), and you may specify a Data Provider value of NONE in the connection string parameter of the [Connection](connection-object-ado.md) object's [Open](open-method-ado-connection.md) method. For more information, see [Required Providers for Data Shaping](required-providers-for-data-shaping.md).</span></span>

```vb
    Dim cn As New ADODB.Connection
    Dim rsCustomers As New ADODB.Recordset
    
    cn.Open "Provider=MSDataShape;Data Provider=NONE;"
     
    strShape = _
    "SHAPE APPEND NEW adInteger AS CustID," & _
                " NEW adChar(25) AS FirstName," & _
                " NEW adChar(25) AS LastName," & _
                " NEW adChar(12) AS SSN," & _
                " NEW adChar(50) AS Address," & _
             " ((SHAPE APPEND NEW adChar(80) AS VIN_NO," & _
                            " NEW adInteger AS CustID," & _
                            " NEW adChar(20) AS BodyColor, " & _
                         " ((SHAPE APPEND NEW adChar(80) AS VIN_NO," & _
                                        " NEW adChar(20) AS Make, " & _
                                        " NEW adChar(20) AS Model," & _
                                        " NEW adChar(4) AS Year) " & _
                            " AS VINS RELATE VIN_NO TO VIN_NO))" & _
                " AS Vehicles RELATE CustID TO CustID) "
     
    rsCustomers.Open strShape, cn, adOpenStatic, adLockOptimistic, -1
```

<span data-ttu-id="7088c-107">Une fois que l' **objet Recordset** a été créé, il peut être rempli, manipulé ou conservé dans un fichier.</span><span class="sxs-lookup"><span data-stu-id="7088c-107">After the **Recordset** has been fabricated, it can be populated, manipulated, or persisted to a file.</span></span>
