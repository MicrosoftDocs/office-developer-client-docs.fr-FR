---
title: Table, objet (ADOX)
TOCTitle: Table Object (ADOX)
ms:assetid: 53a3e2f9-4ec0-8fed-d482-4f995921587b
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249273(v=office.15)
ms:contentKeyID: 48544874
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: fa7ea75110587dc9c32cfe8c2c6c0c1862243ffa
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470552"
---
# <a name="table-object-adox"></a><span data-ttu-id="2761a-102">Table, objet (ADOX)</span><span class="sxs-lookup"><span data-stu-id="2761a-102">Table Object (ADOX)</span></span>


<span data-ttu-id="2761a-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="2761a-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="2761a-104">Représente une table de base de données comprenant des colonnes, des index et des clés.</span><span class="sxs-lookup"><span data-stu-id="2761a-104">Represents a database table including columns, indexes, and keys.</span></span>

## <a name="remarks"></a><span data-ttu-id="2761a-105">Remarques</span><span class="sxs-lookup"><span data-stu-id="2761a-105">Remarks</span></span>

<span data-ttu-id="2761a-106">Le code suivant permet de créer une nouvelle **table**:</span><span class="sxs-lookup"><span data-stu-id="2761a-106">The following code creates a new **Table**:</span></span>

`Dim obj As New Table`

<span data-ttu-id="2761a-107">Avec les propriétés et collections d'un objet **Table**, vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="2761a-107">With the properties and collections of a **Table** object, you can:</span></span>

  - <span data-ttu-id="2761a-108">Identifier la table à l'aide de la propriété [Name](name-property-adox.md).</span><span class="sxs-lookup"><span data-stu-id="2761a-108">Identify the table with the [Name](name-property-adox.md) property.</span></span>

  - <span data-ttu-id="2761a-109">Déterminer le type de la table à l'aide de la propriété [Type](https://msdn.microsoft.com/library/jj250042\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="2761a-109">Determine the type of table with the [Type](https://msdn.microsoft.com/library/jj250042\(v=office.15\)) property.</span></span>

  - <span data-ttu-id="2761a-110">Accéder aux colonnes de base de données de la table à l'aide de la collection [Columns](columns-collection-adox.md).</span><span class="sxs-lookup"><span data-stu-id="2761a-110">Access the database columns of the table with the [Columns](columns-collection-adox.md) collection.</span></span>

  - <span data-ttu-id="2761a-111">Accéder aux index de la table à l'aide de la collection [Indexes](indexes-collection-adox.md).</span><span class="sxs-lookup"><span data-stu-id="2761a-111">Access the indexes of the table with the [Indexes](indexes-collection-adox.md) collection.</span></span>

  - <span data-ttu-id="2761a-112">Accéder aux clés de la table à l'aide de la collection [Keys](keys-collection-adox.md).</span><span class="sxs-lookup"><span data-stu-id="2761a-112">Access the keys of the table with the [Keys](keys-collection-adox.md) collection.</span></span>

  - <span data-ttu-id="2761a-113">Spécifier le [catalogue](catalog-object-adox.md) auquel appartient la table à l'aide de la propriété [ParentCatalog](parentcatalog-property-adox.md).</span><span class="sxs-lookup"><span data-stu-id="2761a-113">Specify the [Catalog](catalog-object-adox.md) that owns the table with the [ParentCatalog](parentcatalog-property-adox.md) property.</span></span>

  - <span data-ttu-id="2761a-114">Renvoyer des informations de date à l'aide des propriétés [DateCreated](datecreated-property-adox.md) et [DateModified](datemodified-property-adox.md).</span><span class="sxs-lookup"><span data-stu-id="2761a-114">Return date information with the [DateCreated](datecreated-property-adox.md) and [DateModified](datemodified-property-adox.md) properties.</span></span>

  - <span data-ttu-id="2761a-115">Accéder aux propriétés de table spécifiques au fournisseur à l'aide de la collection [Properties](properties-collection-ado.md).</span><span class="sxs-lookup"><span data-stu-id="2761a-115">Access provider-specific table properties with the [Properties](properties-collection-ado.md) collection.</span></span>


> [!NOTE]
> <P><span data-ttu-id="2761a-p101">[!REMARQUE] Votre fournisseur de données peut ne pas prendre en charge toutes les propriétés de l'objet <STRONG>Table</STRONG>. Une erreur survient si vous avez défini une valeur pour une propriété non prise en charge par le fournisseur. Pour les nouveaux objets <STRONG>Table</STRONG>, l'erreur survient lorsque l'objet est ajouté à la collection. Pour les objets existants, elle survient au moment de la définition de la propriété.</span><span class="sxs-lookup"><span data-stu-id="2761a-p101">Your data provider may not support all properties of <STRONG>Table</STRONG> objects. An error will occur if you have set a value for a property that the provider does not support. For new <STRONG>Table</STRONG> objects, the error will occur when the object is appended to the collection. For existing objects, the error will occur when setting the property.</span></span></P>



<span data-ttu-id="2761a-p102">Lors de la création des objets **Table**, l'existence d'une valeur par défaut appropriée pour une propriété facultative ne garantit pas que votre fournisseur prend en charge la propriété. Pour plus d'informations sur les propriétés prises en charge par votre fournisseur, reportez-vous à sa documentation.</span><span class="sxs-lookup"><span data-stu-id="2761a-p102">When creating **Table** objects, the existence of an appropriate default value for an optional property does not guarantee that your provider supports the property. For more information about which properties your provider supports, see your provider documentation.</span></span>
