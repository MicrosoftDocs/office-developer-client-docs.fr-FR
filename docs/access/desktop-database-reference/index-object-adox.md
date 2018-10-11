---
title: Index, objet (ADOX)
TOCTitle: Index Object (ADOX)
ms:assetid: fe368ab1-e396-4684-d930-18b0ba58a925
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250304(v=office.15)
ms:contentKeyID: 48548929
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 535ca6a597c6dd580146f6142731897600264526
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470492"
---
# <a name="index-object-adox"></a><span data-ttu-id="8deb9-102">Index, objet (ADOX)</span><span class="sxs-lookup"><span data-stu-id="8deb9-102">Index Object (ADOX)</span></span>

<span data-ttu-id="8deb9-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="8deb9-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="8deb9-104">Représente un index d'une table de base de données.</span><span class="sxs-lookup"><span data-stu-id="8deb9-104">Represents an index from a database table.</span></span>

## <a name="remarks"></a><span data-ttu-id="8deb9-105">Remarques</span><span class="sxs-lookup"><span data-stu-id="8deb9-105">Remarks</span></span>

<span data-ttu-id="8deb9-106">Le code suivant permet de créer un nouvel **index**:</span><span class="sxs-lookup"><span data-stu-id="8deb9-106">The following code creates a new **Index**:</span></span>

`Dim obj As New Index`

<span data-ttu-id="8deb9-107">Avec les propriétés et collections d'un objet **Index**, vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="8deb9-107">With the properties and collections of an **Index** object, you can:</span></span>

  - <span data-ttu-id="8deb9-108">Identifier l'index à l'aide de la propriété [Name](name-property-adox.md).</span><span class="sxs-lookup"><span data-stu-id="8deb9-108">Identify the index with the [Name](name-property-adox.md) property.</span></span>

  - <span data-ttu-id="8deb9-109">Accéder aux colonnes de base de données de l'index à l'aide de la collection [Columns](columns-collection-adox.md).</span><span class="sxs-lookup"><span data-stu-id="8deb9-109">Access the database columns of the index with the [Columns](columns-collection-adox.md) collection.</span></span>

  - <span data-ttu-id="8deb9-110">Spécifier si les clés d'index doivent être uniques à l'aide de la propriété [Unique](unique-property-adox.md).</span><span class="sxs-lookup"><span data-stu-id="8deb9-110">Specify whether the index keys must be unique with the [Unique](unique-property-adox.md) property.</span></span>

  - <span data-ttu-id="8deb9-111">Spécifier si l'index est la clé primaire d'une table à l'aide de la propriété [PrimaryKey](primarykey-property-adox.md).</span><span class="sxs-lookup"><span data-stu-id="8deb9-111">Specify whether the index is the primary key for a table with the [PrimaryKey](primarykey-property-adox.md) property.</span></span>

  - <span data-ttu-id="8deb9-112">Spécifier si les enregistrements dont les champs d'index contiennent des valeurs nulles ont des entrées d'index à l'aide de la propriété [IndexNulls](indexnulls-property-adox.md).</span><span class="sxs-lookup"><span data-stu-id="8deb9-112">Specify whether records that have null values in their index fields have index entries with the [IndexNulls](indexnulls-property-adox.md) property.</span></span>

  - <span data-ttu-id="8deb9-113">Spécifier si l'index est organisé en clusters à l'aide de la propriété [Clustered](clustered-property-adox.md).</span><span class="sxs-lookup"><span data-stu-id="8deb9-113">Specify whether the index is clustered with the [Clustered](clustered-property-adox.md) property.</span></span>

  - <span data-ttu-id="8deb9-114">Accéder aux propriétés d'index spécifiques au fournisseur à l'aide de la collection [Properties](properties-collection-ado.md).</span><span class="sxs-lookup"><span data-stu-id="8deb9-114">Access provider-specific index properties with the [Properties](properties-collection-ado.md) collection.</span></span>


> [!NOTE]
> <span data-ttu-id="8deb9-115">[!REMARQUE] Une erreur survient lors de l'ajout d'une [colonne](column-object-adox.md) à la collection **Columns** d'un **index** si la **colonne** n'existe pas dans un objet [Table](table-object-adox.md) déjà ajouté à la collection [Tables](tables-collection-adox.md).</span><span class="sxs-lookup"><span data-stu-id="8deb9-115">An error will occur when appending a [Column](column-object-adox.md) to the **Columns** collection of an **Index** if the **Column** does not exist in a [Table](table-object-adox.md) object already appended to the [Tables](tables-collection-adox.md) collection.</span></span>

<span data-ttu-id="8deb9-p101">Votre fournisseur de données peut ne pas prendre en charge toutes les propriétés de l'objet **Index**. Une erreur survient si vous avez défini une valeur pour une propriété non prise en charge par le fournisseur. Pour les nouveaux objets **Index**, l'erreur survient lorsque l'objet est ajouté à la collection. Pour les objets existants, elle survient au moment de la définition de la propriété.</span><span class="sxs-lookup"><span data-stu-id="8deb9-p101">Your data provider may not support all properties of **Index** objects. An error will occur if you have set a value for a property that is not supported by the provider. For new **Index** objects, the error will occur when the object is appended to the collection. For existing objects, the error will occur when setting the property.</span></span>

<span data-ttu-id="8deb9-p102">Lors de la création des objets **Index**, l'existence d'une valeur par défaut appropriée pour une propriété facultative ne garantit pas que votre fournisseur prend en charge la propriété. Pour plus d'informations sur les propriétés prises en charge par votre fournisseur, reportez-vous à sa documentation.</span><span class="sxs-lookup"><span data-stu-id="8deb9-p102">When creating **Index** objects, the existence of an appropriate default value for an optional property does not guarantee that your provider supports the property. For more information about which properties your provider supports, see your provider documentation.</span></span>
