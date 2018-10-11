---
title: Append, méthode (ADO)
TOCTitle: Append Method (ADO)
ms:assetid: cca133af-2b95-877d-0488-0d99631623f2
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250014(v=office.15)
ms:contentKeyID: 48547742
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: b47b7b0514b78a89425e47962c36b092e35677ea
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471291"
---
# <a name="append-method-ado"></a><span data-ttu-id="7d48e-102">Append, méthode (ADO)</span><span class="sxs-lookup"><span data-stu-id="7d48e-102">Append Method (ADO)</span></span>


<span data-ttu-id="7d48e-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="7d48e-103">**Applies to**: Access 2013 | Office 2013</span></span>



<span data-ttu-id="7d48e-p101">Ajoute un objet à une collection. S'il s'agit de la collection [Fields](fields-collection-ado.md), un nouvel objet [Field](field-object-ado.md) peut être créé avant d'être ajouté à la collection.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p101">Appends an object to a collection. If the collection is [Fields](fields-collection-ado.md), a new [Field](field-object-ado.md) object may be created before it is appended to the collection.</span></span>

## <a name="syntax"></a><span data-ttu-id="7d48e-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="7d48e-106">Syntax</span></span>

<span data-ttu-id="7d48e-107">la *collection*. Ajouter *l’objet*</span><span class="sxs-lookup"><span data-stu-id="7d48e-107">*collection*.Append *object*</span></span>

<span data-ttu-id="7d48e-108">*champs*. Ajouter le *nom*, *Type*, *DefinedSize*, *attributs*, *FieldValue*</span><span class="sxs-lookup"><span data-stu-id="7d48e-108">*fields*.Append *Name*, *Type*, *DefinedSize*, *Attrib*, *FieldValue*</span></span>

## <a name="parameters"></a><span data-ttu-id="7d48e-109">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7d48e-109">Parameters</span></span>

  - <span data-ttu-id="7d48e-110">*collection*</span><span class="sxs-lookup"><span data-stu-id="7d48e-110">*collection*</span></span>

  - <span data-ttu-id="7d48e-111">Objet Collection.</span><span class="sxs-lookup"><span data-stu-id="7d48e-111">A collection object.</span></span>

  - <span data-ttu-id="7d48e-112">*fields*</span><span class="sxs-lookup"><span data-stu-id="7d48e-112">*fields*</span></span>

  - <span data-ttu-id="7d48e-113">Collection **Fields**.</span><span class="sxs-lookup"><span data-stu-id="7d48e-113">A **Fields** collection.</span></span>

  - <span data-ttu-id="7d48e-114">*object*</span><span class="sxs-lookup"><span data-stu-id="7d48e-114">*object*</span></span>

  - <span data-ttu-id="7d48e-115">Variable objet représentant l'objet à ajouter.</span><span class="sxs-lookup"><span data-stu-id="7d48e-115">An object variable that represents the object to be appended.</span></span>

  - <span data-ttu-id="7d48e-116">*Name*</span><span class="sxs-lookup"><span data-stu-id="7d48e-116">*Name*</span></span>

  - <span data-ttu-id="7d48e-117">Valeur **String** contenant le nom du nouvel objet **Field**. Ce nom doit être différent de ceux des autres objets de la collection *Fields*.</span><span class="sxs-lookup"><span data-stu-id="7d48e-117">A **String** value that contains the name of the new **Field** object, and must not be the same name as any other object in *fields*.</span></span>

  - <span data-ttu-id="7d48e-118">*Type*</span><span class="sxs-lookup"><span data-stu-id="7d48e-118">*Type*</span></span>

  - <span data-ttu-id="7d48e-p102">Valeur [DataTypeEnum](datatypeenum.md) qui, par défaut, est définie à **adEmpty** et qui spécifie le type de données du nouveau champ. Les types de données suivants ne sont pas pris en charge par ADO et ne doivent pas être utilisés lors de l'ajout de nouveaux champs à un objet **Recordset**: **adIDispatch**, **adIUnknown**, **adVariant**.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p102">A [DataTypeEnum](datatypeenum.md) value, whose default value is **adEmpty**, that specifies the data type of the new field. The following data types are not supported by ADO, and should not be used when appending new fields to a **Recordset**: **adIDispatch**, **adIUnknown**, **adVariant**.</span></span>

  - <span data-ttu-id="7d48e-121">*DefinedSize*</span><span class="sxs-lookup"><span data-stu-id="7d48e-121">*DefinedSize*</span></span>

  - <span data-ttu-id="7d48e-p103">Facultatif. Valeur de type **Long** représentant la taille définie, en caractères ou en octets, du nouveau champ. La valeur par défaut de ce paramètre est dérivée de *Type*. Les champs dont la valeur DefinedSize est supérieure à 255 octets sont traités comme des colonnes de longueur variable. (La valeur *DefinedSize* par défaut n'est pas spécifiée.)</span><span class="sxs-lookup"><span data-stu-id="7d48e-p103">Optional. A **Long** value that represents the defined size, in characters or bytes, of the new field. The default value for this parameter is derived from *Type*. Fields with a DefinedSize greater than 255 bytes, and treated as variable length columns. (The default *DefinedSize* is unspecified.)</span></span>

  - <span data-ttu-id="7d48e-127">*Attrib*</span><span class="sxs-lookup"><span data-stu-id="7d48e-127">*Attrib*</span></span>

  - <span data-ttu-id="7d48e-p104">Facultatif. Valeur [FieldAttributeEnum](fieldattributeenum.md), définie par défaut à **adFldDefault**, qui spécifie les attributs du nouveau champ. Si cette valeur n’est pas spécifiée, le champ contient les attributs dérivés de *Type*.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p104">Optional. A [FieldAttributeEnum](fieldattributeenum.md) value, whose default value is **adFldDefault**, that specifies attributes for the new field. If this value is not specified, the field will contain attributes derived from *Type*.</span></span>

  - <span data-ttu-id="7d48e-131">*FieldValue*</span><span class="sxs-lookup"><span data-stu-id="7d48e-131">*FieldValue*</span></span>

  - <span data-ttu-id="7d48e-p105">Facultatif. **Variant** représentant la valeur du nouveau champ. S'il n'est pas spécifié, le champ est ajouté avec la valeur Null.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p105">Optional. A **Variant** that represents the value for the new field. If not specified, then the field is appended with a null value.</span></span>

## <a name="remarks"></a><span data-ttu-id="7d48e-135">Notes</span><span class="sxs-lookup"><span data-stu-id="7d48e-135">Remarks</span></span>

<span data-ttu-id="7d48e-136">**Collection Parameters**</span><span class="sxs-lookup"><span data-stu-id="7d48e-136">**Parameters Collection**</span></span>

<span data-ttu-id="7d48e-p106">Vous devez définir la propriété [Type](type-property-ado.md) d'un objet [Parameter](parameter-object-ado.md) avant de l'ajouter à la collection **Parameters**. Si vous sélectionnez un type de données de longueur variable, vous devez également attribuer une valeur supérieure à zéro à la propriété [Size](size-property-ado.md).</span><span class="sxs-lookup"><span data-stu-id="7d48e-p106">You must set the [Type](type-property-ado.md) property of a [Parameter](parameter-object-ado.md) object before appending it to the **Parameters** collection. If you select a variable-length data type, you must also set the [Size](size-property-ado.md) property to a value greater than zero.</span></span>

<span data-ttu-id="7d48e-p107">Lorsque vous décrivez vous-même des paramètres, vous limitez le nombre d'appels du fournisseur et vous obtenez de meilleures performances lorsque vous utilisez des procédures stockées ou des requêtes paramétrées. Toutefois, vous devez connaître les propriétés des paramètres associés à la procédure stockée ou à la requête paramétrée que vous voulez appeler.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p107">Describing parameters yourself minimizes calls to the provider and consequently improves performance when using stored procedures or parameterized queries. However, you must know the properties of the parameters associated with the stored procedure or parameterized query that you want to call.</span></span>

<span data-ttu-id="7d48e-p108">Faites appel à la méthode [CreateParameter](createparameter-method-ado.md) pour créer des objets **Parameter** dotés des paramètres de propriété appropriés et la méthode **Append** pour les ajouter à la collection [Parameters](parameters-collection-ado.md). Vous pouvez ainsi définir et renvoyer des valeurs de paramètres sans avoir à appeler le fournisseur pour obtenir les informations appropriées. Si vous faites appel à un fournisseur qui ne fournit pas de données de paramètre, vous devez utiliser cette méthode pour remplir manuellement la collection **Parameters** sans quoi vous ne serez pas en mesure d'employer un quelconque paramètre.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p108">Use the [CreateParameter](createparameter-method-ado.md) method to create **Parameter** objects with the appropriate property settings and use the **Append** method to add them to the [Parameters](parameters-collection-ado.md) collection. This lets you set and return parameter values without having to call the provider for the parameter information. If you are writing to a provider that does not supply parameter information, you must use this method to manually populate the **Parameters** collection in order to use parameters at all.</span></span>

<span data-ttu-id="7d48e-144">**Collection Fields**</span><span class="sxs-lookup"><span data-stu-id="7d48e-144">**Fields Collection**</span></span>

<span data-ttu-id="7d48e-145">Le paramètre *FieldValue* n’est valide que lorsque vous ajoutez un objet **Field** à un objet [Record](record-object-ado.md) , pas à un objet **Recordset** .</span><span class="sxs-lookup"><span data-stu-id="7d48e-145">The *FieldValue* parameter is only valid when adding a **Field** object to a [Record](record-object-ado.md) object, not to a **Recordset** object.</span></span> <span data-ttu-id="7d48e-146">Avec un objet **Record** , vous pouvez ajouter des champs et fournir des valeurs en même temps.</span><span class="sxs-lookup"><span data-stu-id="7d48e-146">With a **Record** object, you may append fields and provide values at the same time.</span></span> <span data-ttu-id="7d48e-147">Avec un objet **Recordset** , vous devez créer des champs lorsque le **Recordset** est fermé, puis ouvrir le **jeu d’enregistrements** et affecter des valeurs aux champs.</span><span class="sxs-lookup"><span data-stu-id="7d48e-147">With a **Recordset** object, you must create fields while the **Recordset** is closed, then open the **Recordset** and assign values to the fields.</span></span>


> [!NOTE]
> <P><span data-ttu-id="7d48e-p110">[!REMARQUE] Pour les nouveaux objets <STRONG>Field</STRONG> qui ont été ajoutés à la collection <STRONG>Fields</STRONG> d'un objet <STRONG>Record</STRONG>, la propriété <A href="value-property-ado.md">Value</A> doit être définie pour qu'une propriété <STRONG>Field</STRONG> puisse être spécifiée. Une valeur spécifique doit avoir été affectée au préalable à la propriété <STRONG>Value</STRONG> et la méthode <A href="update-method-ado.md">Update</A> doit avoir été appelée sur la collection <STRONG>Fields</STRONG>. Vous pouvez alors accéder à d'autres propriétés comme <A href="type-property-ado.md">Type</A> ou <A href="attributes-property-ado.md">Attributes</A>.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p110">For new <STRONG>Field</STRONG> objects that have been appended to the <STRONG>Fields</STRONG> collection of a <STRONG>Record</STRONG> object, the <A href="value-property-ado.md">Value</A> property must be set before any other <STRONG>Field</STRONG> properties can be specified. First, a specific value for the <STRONG>Value</STRONG> property must have been assigned and <A href="update-method-ado.md">Update</A> on the <STRONG>Fields</STRONG> collection called. Then, other properties such as <A href="type-property-ado.md">Type</A> or <A href="attributes-property-ado.md">Attributes</A> can be accessed.</span></span></P>



<span data-ttu-id="7d48e-p111">Une erreur se produit si vous tentez d'ajouter des objets **Field** ayant les types de données suivants (**DataTypeEnum**) à la collection **Fields** : **adArray**, **adChapter**, **adEmpty**, **adPropVariant** et **adUserDefined**. De même, les types de données suivants ne sont pas pris en charge par ADO : **adIDispatch**, **adIUnknown** et **adIVariant**. L'ajout de ces derniers ne génère pas d'erreur mais vous risquez d'obtenir des résultats imprévisibles et notamment des pertes de mémoire lorsque vous les utilisez.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p111">**Field** objects of the following data types (**DataTypeEnum**) cannot be appended to the **Fields** collection and will cause an error to occur: **adArray**, **adChapter**, **adEmpty**, **adPropVariant**, and **adUserDefined**. Also, the following data types are not supported by ADO: **adIDispatch**, **adIUnknown**, and **adIVariant**. For these types, no error will occur when appended, but usage can produce unpredictable results including memory leaks.</span></span>

<span data-ttu-id="7d48e-154">**Objet Recordset**</span><span class="sxs-lookup"><span data-stu-id="7d48e-154">**Recordset**</span></span>

<span data-ttu-id="7d48e-155">Si vous ne définissez pas la propriété [CursorLocation](cursorlocation-property-ado.md) avant d'appeler la méthode **Append**, **CursorLocation** prend automatiquement la valeur **adUseClient** (une valeur [CursorLocationEnum](cursorlocationenum.md)) lorsque la méthode [Open](recordset-object-ado.md) de l'objet [Recordset](open-method-ado-recordset.md) est appelée.</span><span class="sxs-lookup"><span data-stu-id="7d48e-155">If you do not set the [CursorLocation](cursorlocation-property-ado.md) property before calling the **Append** method, **CursorLocation** will be set to **adUseClient** (a [CursorLocationEnum](cursorlocationenum.md) value) automatically when the [Recordset](recordset-object-ado.md) object's [Open](open-method-ado-recordset.md) method is called.</span></span>

<span data-ttu-id="7d48e-p112">Vous obtenez une erreur d'exécution lorsque la méthode **Append** est appelée sur la collection **Fields** d'un objet **Recordset** ouvert ou d'un objet **Recordset** dont la propriété [ActiveConnection](activeconnection-property-ado.md) a été définie. Vous ne pouvez ajouter des champs qu'à un objet **Recordset** qui est fermé et qui n'a pas encore été connecté à une source de données. Ceci se produit, en principe, lorsqu'un objet **Recordset** est créé à l'aide d'une méthode [CreateRecordset](createrecordset-method-rds.md) ou qu'il est assigné à une variable objet.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p112">A run-time error will occur if the **Append** method is called on the **Fields** collection of an open **Recordset**, or on a **Recordset** where the [ActiveConnection](activeconnection-property-ado.md) property has been set. You can only append fields to a **Recordset** that is not open and has not yet been connected to a data source. This is typically the case when a **Recordset** object is fabricated with the [CreateRecordset](createrecordset-method-rds.md) method or assigned to an object variable.</span></span>

<span data-ttu-id="7d48e-159">**Objet Record**</span><span class="sxs-lookup"><span data-stu-id="7d48e-159">**Record**</span></span>

<span data-ttu-id="7d48e-p113">Vous n'obtenez pas d'erreur d'exécution lorsque la méthode **Append** est appelée sur la collection **Fields** d'un objet **Record** ouvert. Le nouveau champ est ajouté à la collection **Fields** de l'objet **Record**. Si l'objet **Record** est dérivé d'un objet **Recordset**, le nouveau champ n'apparaît pas dans la collection **Fields** de l'objet **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p113">A run-time error will not occur if the **Append** method is called on the **Fields** collection of an open **Record**. The new field will be added to the **Record** object's **Fields** collection. If the **Record** was derived from a **Recordset**, then the new field will not appear in the **Recordset** object's **Fields** collection.</span></span>

<span data-ttu-id="7d48e-p114">Vous pouvez créer un champ non existant et l'ajouter à la collection **Fields** en affectant une valeur à ce champ comme s'il existait déjà dans la collection. Cette affectation déclenche la création et l'ajout automatiques de l'objet **Field**, puis son exécution effective.</span><span class="sxs-lookup"><span data-stu-id="7d48e-p114">A non-existent field can be created and appended to the **Fields** collection by assigning a value to the field object as if it already existed in the collection. The assignment will trigger the automatic creation and appending of the **Field** object, then the assignment will be completed.</span></span>

<span data-ttu-id="7d48e-165">Une fois l'objet **Field** ajouté à la collection **Fields** d'un objet **Record**, appelez la méthode **Update** de la collection **Fields** pour enregistrer cette modification.</span><span class="sxs-lookup"><span data-stu-id="7d48e-165">After appending a **Field** to a **Record** object's **Fields** collection, call the **Update** method of the **Fields** collection to save the change.</span></span>
