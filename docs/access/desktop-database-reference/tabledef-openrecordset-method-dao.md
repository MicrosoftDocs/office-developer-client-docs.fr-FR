---
title: TableDef.OpenRecordset Method (DAO)
TOCTitle: OpenRecordset Method
ms:assetid: f4c9c89c-3348-d3c9-ce76-dd11e5ee11a7
ms:mtpsurl: https://msdn.microsoft.com/library/Ff836703(v=office.15)
ms:contentKeyID: 48548696
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: d7b9b6ccac186f87d0a22a4a55108594cfc55ac9
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470221"
---
# <a name="tabledefopenrecordset-method-dao"></a><span data-ttu-id="d8ffc-102">TableDef.OpenRecordset Method (DAO)</span><span class="sxs-lookup"><span data-stu-id="d8ffc-102">TableDef.OpenRecordset Method (DAO)</span></span>


<span data-ttu-id="d8ffc-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="d8ffc-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="d8ffc-104">Crée un objet **[Recordset](recordset-object-dao.md)** et l'ajoute à la collection **Recordsets**.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-104">Creates a new **[Recordset](recordset-object-dao.md)** object and appends it to the **Recordsets** collection.</span></span>

## <a name="syntax"></a><span data-ttu-id="d8ffc-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="d8ffc-105">Syntax</span></span>

<span data-ttu-id="d8ffc-106">*expression* . OpenRecordset (***Type***, ***Options***)</span><span class="sxs-lookup"><span data-stu-id="d8ffc-106">*expression* .OpenRecordset(***Type***, ***Options***)</span></span>

<span data-ttu-id="d8ffc-107">*expression* Variable qui représente un objet **TableDef** .</span><span class="sxs-lookup"><span data-stu-id="d8ffc-107">*expression* A variable that represents a **TableDef** object.</span></span>

### <a name="parameters"></a><span data-ttu-id="d8ffc-108">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d8ffc-108">Parameters</span></span>

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="d8ffc-109">Name</span><span class="sxs-lookup"><span data-stu-id="d8ffc-109">Name</span></span></p></th>
<th><p><span data-ttu-id="d8ffc-110">Obligatoire/Facultatif</span><span class="sxs-lookup"><span data-stu-id="d8ffc-110">Required/Optional</span></span></p></th>
<th><p><span data-ttu-id="d8ffc-111">Type de données</span><span class="sxs-lookup"><span data-stu-id="d8ffc-111">Data Type</span></span></p></th>
<th><p><span data-ttu-id="d8ffc-112">Description</span><span class="sxs-lookup"><span data-stu-id="d8ffc-112">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="d8ffc-113">Name</span><span class="sxs-lookup"><span data-stu-id="d8ffc-113">Name</span></span></p></td>
<td><p><span data-ttu-id="d8ffc-114">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="d8ffc-114">Required</span></span></p></td>
<td><p><span data-ttu-id="d8ffc-115"><strong>Chaîne</strong></span><span class="sxs-lookup"><span data-stu-id="d8ffc-115"><strong>String</strong></span></span></p></td>
<td><p><span data-ttu-id="d8ffc-p101">Source des enregistrements du nouveau <strong>Recordset</strong>. La source peut être un nom de table, un nom de requête ou une instruction SQL qui renvoie des enregistrements. Pour les objets <strong>Recordset</strong> de type table dans les bases de données du moteur de base de données Microsoft Access, la source peut uniquement être un nom de table.  </span><span class="sxs-lookup"><span data-stu-id="d8ffc-p101">The source of the records for the new <strong>Recordset</strong>. The source can be a table name, a query name, or an SQL statement that returns records. For table-type <strong>Recordset</strong> objects in Microsoft Access database engine databases, the source can only be a table name.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="d8ffc-119">Type</span><span class="sxs-lookup"><span data-stu-id="d8ffc-119">Type</span></span></p></td>
<td><p><span data-ttu-id="d8ffc-120">Facultatif</span><span class="sxs-lookup"><span data-stu-id="d8ffc-120">Optional</span></span></p></td>
<td><p><span data-ttu-id="d8ffc-121"><strong>Variante</strong></span><span class="sxs-lookup"><span data-stu-id="d8ffc-121"><strong>Variant</strong></span></span></p></td>
<td><p><span data-ttu-id="d8ffc-122">Constante <strong><a href="recordsettypeenum-enumeration-dao.md">RecordsetTypeEnum</a></strong> qui indique le type de <strong>Recordset</strong> à ouvrir.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-122">A <strong><a href="recordsettypeenum-enumeration-dao.md">RecordsetTypeEnum</a></strong> constant that indicates the type of <strong>Recordset</strong> to open.</span></span></p>

> [!NOTE]
> <P><span data-ttu-id="d8ffc-p102">Si vous ouvrez un <STRONG>Recordset</STRONG> dans un espace de travail Microsoft Access et que vous n’indiquez aucun type, <STRONG>OpenRecordset</STRONG> crée un <STRONG>Recordset</STRONG> de type table, si possible. Si vous spécifiez une table liée ou une requête, <STRONG>OpenRecordset</STRONG> crée un <STRONG>Recordset</STRONG> de type feuille de réponse dynamique.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-p102">If you open a <STRONG>Recordset</STRONG> in a Microsoft Access workspace and you don't specify a type, <STRONG>OpenRecordset</STRONG> creates a table-type <STRONG>Recordset</STRONG>, if possible. If you specify a linked table or query, <STRONG>OpenRecordset</STRONG> creates a dynaset-type <STRONG>Recordset</STRONG>.</span></span></P>


</td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="d8ffc-125">Options</span><span class="sxs-lookup"><span data-stu-id="d8ffc-125">Options</span></span></p></td>
<td><p><span data-ttu-id="d8ffc-126">Facultatif</span><span class="sxs-lookup"><span data-stu-id="d8ffc-126">Optional</span></span></p></td>
<td><p><span data-ttu-id="d8ffc-127"><strong>Variant</strong></span><span class="sxs-lookup"><span data-stu-id="d8ffc-127"><strong>Variant</strong></span></span></p></td>
<td><p><span data-ttu-id="d8ffc-128">Combinaison de constantes <strong><a href="recordsetoptionenum-enumeration-dao.md">RecordsetOptionEnum</a></strong> qui indiquent les caractéristiques du nouveau <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-128">A combination of <strong><a href="recordsetoptionenum-enumeration-dao.md">RecordsetOptionEnum</a></strong> constants that specify characteristics of the new <strong>Recordset</strong>.</span></span></p>

> [!NOTE]
> <P><span data-ttu-id="d8ffc-p103">Les constantes <STRONG>dbConsistent</STRONG> et <STRONG>dbInconsistent</STRONG> s’excluent mutuellement, et l’utilisation de ces deux constantes peut entraîner une erreur. L’indication d’un argument lockedits lorsque options utilise la constante <STRONG>dbReadOnly</STRONG> génère également une erreur.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-p103">The constants <STRONG>dbConsistent</STRONG> and <STRONG>dbInconsistent</STRONG> are mutually exclusive, and using both causes an error. Supplying a lockedits argument when options uses the <STRONG>dbReadOnly</STRONG> constant also causes an error.</span></span></P>


</td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="d8ffc-131">VerrouillerModification</span><span class="sxs-lookup"><span data-stu-id="d8ffc-131">LockEdit</span></span></p></td>
<td><p><span data-ttu-id="d8ffc-132">Facultatif</span><span class="sxs-lookup"><span data-stu-id="d8ffc-132">Optional</span></span></p></td>
<td><p><span data-ttu-id="d8ffc-133"><strong>Variant</strong></span><span class="sxs-lookup"><span data-stu-id="d8ffc-133"><strong>Variant</strong></span></span></p></td>
<td><p><span data-ttu-id="d8ffc-134">Constante <strong><a href="locktypeenum-enumeration-dao.md">LockTypeEnum</a></strong> qui détermine le verrouillage du <strong>Recordset</strong>.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-134">A <strong><a href="locktypeenum-enumeration-dao.md">LockTypeEnum</a></strong> constant that determines the locking for the <strong>Recordset</strong>.</span></span></p>

> [!NOTE]
> <P><span data-ttu-id="d8ffc-p104">Vous pouvez utiliser <STRONG>dbReadOnly</STRONG> dans l’argument options ou l’argument lockededits, mais pas dans les deux. Si vous l’utilisez pour les deux arguments, une erreur d’exécution se produit.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-p104">You can use <STRONG>dbReadOnly</STRONG> in either the options argument or the lockedits argument, but not both. If you use it for both arguments, a run-time error occurs.</span></span></P>


</td>
</tr>
</tbody>
</table>


### <a name="return-value"></a><span data-ttu-id="d8ffc-137">Valeur renvoyée</span><span class="sxs-lookup"><span data-stu-id="d8ffc-137">Return Value</span></span>

<span data-ttu-id="d8ffc-138">Recordset</span><span class="sxs-lookup"><span data-stu-id="d8ffc-138">Recordset</span></span>

## <a name="remarks"></a><span data-ttu-id="d8ffc-139">Remarques</span><span class="sxs-lookup"><span data-stu-id="d8ffc-139">Remarks</span></span>

<span data-ttu-id="d8ffc-p105">En règle générale, si l'utilisateur obtient cette erreur pendant la mise à jour d'un enregistrement, votre code doit actualiser le contenu des champs et extraire les valeurs nouvellement modifiées. Si l'erreur se produit lors de la suppression d'un enregistrement, votre code doit présenter les nouvelles données de l'enregistrement à l'utilisateur et afficher un message indiquant que les données ont été récemment modifiées. À ce stade, votre code peut demander à l'utilisateur de confirmer sa volonté de supprimer l'enregistrement.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-p105">Typically, if the user gets this error while updating a record, your code should refresh the contents of the fields and retrieve the newly modified values. If the error occurs while deleting a record, your code could display the new record data to the user and a message indicating that the data has recently changed. At this point, your code can request a confirmation that the user still wants to delete the record.</span></span>

<span data-ttu-id="d8ffc-143">Par ailleurs, il est recommandé d'utiliser la constante **dbSeeChanges** si vous ouvrez un objet **Recordset** dans un espace de travail ODBC connecté au moteur de base de données Microsoft Access pour une table Microsoft SQL Server 6.0 (ou version ultérieure) qui comporte une colonne IDENTITY. À défaut, une erreur risque de se produire.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-143">You should also use the **dbSeeChanges** constant if you open a **Recordset** in a Microsoft Access database engine-connected ODBC workspace against a Microsoft SQL Server 6.0 (or later) table that has an IDENTITY column, otherwise an error may result.</span></span>

<span data-ttu-id="d8ffc-p106">L'ouverture de plusieurs objets **Recordset** dans une source de données ODBC peut entraîner un échec car la connexion est occupée par un appel d'objet **OpenRecordset** préalable. Pour contourner ce problème, vous pouvez renseigner totalement l'objet **Recordset** à l'aide de la méthode **MoveLast** dès que l'objet **Recordset** est ouvert.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-p106">Opening more than one **Recordset** on an ODBC data source may fail because the connection is busy with a prior **OpenRecordset** call. One way around this is to fully populate the **Recordset** by using the **MoveLast** method as soon as the **Recordset** is opened.</span></span>

<span data-ttu-id="d8ffc-146">Le fait de fermer un objet **Recordset** par le biais de la méthode **[Close](connection-close-method-dao.md)** entraîne sa suppression automatique au niveau de la collection **Recordsets**.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-146">Closing a **Recordset** with the **[Close](connection-close-method-dao.md)** method automatically deletes it from the **Recordsets** collection.</span></span>


> [!NOTE]
> <P><span data-ttu-id="d8ffc-147">Si <EM>la source</EM> fait référence à une instruction SQL composée d’une chaîne concaténée avec une valeur non entière, et les paramètres système spécifient un caractère décimal américain comme une virgule (par exemple, strSQL = « prix &gt; » &amp; lngPrice et lngPrice = 125,50), une erreur se produit lorsque vous essayez d’ouvrir le <STRONG>jeu d’enregistrements</STRONG>.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-147">If <EM>source</EM> refers to an SQL statement composed of a string concatenated with a non-integer value, and the system parameters specify a non-U.S. decimal character such as a comma (for example, strSQL = "PRICE &gt; " &amp; lngPrice, and lngPrice = 125,50), an error occurs when you try to open the <STRONG>Recordset</STRONG>.</span></span> <span data-ttu-id="d8ffc-148">Cela est dû au fait que pendant la concaténation, le nombre est converti en chaîne en utilisant le caractère décimal par défaut de votre système, et SQL n'accepte que les caractères décimaux usités aux États-Unis.</span><span class="sxs-lookup"><span data-stu-id="d8ffc-148">This is because during concatenation, the number will be converted to a string using your system's default decimal character, and SQL only accepts U.S. decimal characters.</span></span></P>

