---
title: SchemaEnum (référence de base de données du bureau Access)
TOCTitle: SchemaEnum
ms:assetid: 6147b682-3c4f-ea91-fff6-ac73107d206d
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249359(v=office.15)
ms:contentKeyID: 48545208
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: e928271d61cf38232014db1c67482ca6d47fd863
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472130"
---
# <a name="schemaenum"></a><span data-ttu-id="3975c-102">SchemaEnum</span><span class="sxs-lookup"><span data-stu-id="3975c-102">SchemaEnum</span></span>


<span data-ttu-id="3975c-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="3975c-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="3975c-104">Spécifie le type de **Recordset** de schéma extrait par la méthode [OpenSchema](openschema-method-ado.md).</span><span class="sxs-lookup"><span data-stu-id="3975c-104">Specifies the type of schema **Recordset** that the [OpenSchema](openschema-method-ado.md) method retrieves.</span></span>

<span data-ttu-id="3975c-105">**Remarques**</span><span class="sxs-lookup"><span data-stu-id="3975c-105">**Remarks**</span></span>

<span data-ttu-id="3975c-p101">D'autres informations sur la fonction et les colonnes renvoyées pour chaque constante ADO sont présentes dans les rubriques de l'Annexe B du manuel *OLE DB Programmers Reference*. Le nom de chaque rubrique est donné entre parenthèses dans la colonne Description du tableau ci-après.</span><span class="sxs-lookup"><span data-stu-id="3975c-p101">Additional information about the function and columns returned for each ADO constant can be found in topics of Appendix B of the *OLE DB Programmers Reference*. The name of each topic is listed in parentheses in the Description section of the table below.</span></span>

<span data-ttu-id="3975c-108">Vous trouverez plus d’informations sur la fonction et les colonnes renvoyées pour chaque constante ADO MD dans les rubriques du chapitre 23 de la documentation *OLE DB pour OLAP* .</span><span class="sxs-lookup"><span data-stu-id="3975c-108">Additional information about the function and columns returned for each ADO MD constant can be found in topics of Chapter 23 of the *OLE DB for OLAP* documentation.</span></span> <span data-ttu-id="3975c-109">Le nom de chaque rubrique est indiqué entre parenthèses et marqué avec un astérisque (\*) dans la colonne Description du tableau ci-après.</span><span class="sxs-lookup"><span data-stu-id="3975c-109">The name of each topic is listed in parentheses and marked with an asterisk (\*) in the Description column of the table below.</span></span>

<span data-ttu-id="3975c-110">Convertissez les types de données des colonnes de la documentation OLE DB en types de données ADO en vous reportant à la colonne de la rubrique [DataTypeEnum](datatypeenum.md) ADO.</span><span class="sxs-lookup"><span data-stu-id="3975c-110">Translate the data types of columns in the OLE DB documentation to ADO data types by referring to the Description column of the ADO [DataTypeEnum](datatypeenum.md) topic.</span></span> <span data-ttu-id="3975c-111">Par exemple, un type de données OLE DB de **DBTYPE\_WSTR** équivaut à un type de données ADO de **adWChar**.</span><span class="sxs-lookup"><span data-stu-id="3975c-111">For example, an OLE DB data type of **DBTYPE\_WSTR** is equivalent to an ADO data type of **adWChar**.</span></span>

<span data-ttu-id="3975c-p104">ADO génère des résultats de type schéma pour les constantes **adSchemaDBInfoKeywords** et **adSchemaDBInfoLiterals**. ADO crée un **Recordset**, puis renseigne chaque ligne avec les valeurs renvoyées respectivement par les méthodes **IDBInfo::GetKeywords** et **IDBInfo::GetLiteralInfo**. Pour plus d'informations sur ces méthodes, consultez la section IDBInfo du manuel *OLE DB Programmer's Reference.*</span><span class="sxs-lookup"><span data-stu-id="3975c-p104">ADO generates schema-like results for the constants, **adSchemaDBInfoKeywords** and **adSchemaDBInfoLiterals**. ADO creates a **Recordset**, then fills each row with the values returned respectively by the **IDBInfo::GetKeywords** and **IDBInfo::GetLiteralInfo** methods. Additional information about these methods can be found in the IDBInfo section of the *OLE DB Programmer's Reference.*</span></span>

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="3975c-115">Constant</span><span class="sxs-lookup"><span data-stu-id="3975c-115">Constant</span></span></p></th>
<th><p><span data-ttu-id="3975c-116">Valeur</span><span class="sxs-lookup"><span data-stu-id="3975c-116">Value</span></span></p></th>
<th><p><span data-ttu-id="3975c-117">Description</span><span class="sxs-lookup"><span data-stu-id="3975c-117">Description</span></span></p></th>
<th><p><span data-ttu-id="3975c-118">Contraintes de colonne</span><span class="sxs-lookup"><span data-stu-id="3975c-118">Constraint Columns</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="3975c-119"><strong>adSchemaAsserts</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-119"><strong>adSchemaAsserts</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-120">0</span><span class="sxs-lookup"><span data-stu-id="3975c-120">0</span></span></p></td>
<td><p><span data-ttu-id="3975c-121">Renvoie les assertions définies dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-121">Returns the assertions defined in the catalog that are owned by a given user.</span></span> <span data-ttu-id="3975c-122">(ASSERTIONS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-122">(ASSERTIONS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-123">CONSTRAINT_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-123">CONSTRAINT_CATALOG</span></span><br />
<span data-ttu-id="3975c-124">CONSTRAINT_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-124">CONSTRAINT_SCHEMA</span></span><br />
<span data-ttu-id="3975c-125">NOM_CONTRAINTE</span><span class="sxs-lookup"><span data-stu-id="3975c-125">CONSTRAINT_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-126"><strong>adSchemaCatalogs</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-126"><strong>adSchemaCatalogs</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-127">1</span><span class="sxs-lookup"><span data-stu-id="3975c-127">1</span></span></p></td>
<td><p><span data-ttu-id="3975c-128">Renvoie les attributs physiques associés aux catalogues accessibles depuis le DBMS.
</span><span class="sxs-lookup"><span data-stu-id="3975c-128">Returns the physical attributes associated with catalogs accessible from the DBMS.</span></span> <span data-ttu-id="3975c-129">(CATALOGS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-129">(CATALOGS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-130">CATALOG_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-130">CATALOG_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-131"><strong>adSchemaCharacterSets</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-131"><strong>adSchemaCharacterSets</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-132">2</span><span class="sxs-lookup"><span data-stu-id="3975c-132">2</span></span></p></td>
<td><p><span data-ttu-id="3975c-133">Renvoie les jeux de caractères définis dans le catalogue et accessibles par un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-133">Returns the character sets defined in the catalog that are accessible to a given user.</span></span> <span data-ttu-id="3975c-134">(CHARACTER_SETS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-134">(CHARACTER_SETS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-135">CHARACTER_SET_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-135">CHARACTER_SET_CATALOG</span></span><br />
<span data-ttu-id="3975c-136">CHARACTER_SET_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-136">CHARACTER_SET_SCHEMA</span></span><br />
<span data-ttu-id="3975c-137">CHARACTER_SET_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-137">CHARACTER_SET_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-138"><strong>adSchemaCheckConstraints</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-138"><strong>adSchemaCheckConstraints</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-139">5</span><span class="sxs-lookup"><span data-stu-id="3975c-139">5</span></span></p></td>
<td><p><span data-ttu-id="3975c-140">Renvoie les contraintes de contrôle définies dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-140">Returns the check constraints defined in the catalog that are owned by a given user.</span></span> <span data-ttu-id="3975c-141">(CHECK_CONSTRAINTS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-141">(CHECK_CONSTRAINTS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-142">CONSTRAINT_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-142">CONSTRAINT_CATALOG</span></span><br />
<span data-ttu-id="3975c-143">CONSTRAINT_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-143">CONSTRAINT_SCHEMA</span></span><br />
<span data-ttu-id="3975c-144">NOM_CONTRAINTE</span><span class="sxs-lookup"><span data-stu-id="3975c-144">CONSTRAINT_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-145"><strong>adSchemaCollations</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-145"><strong>adSchemaCollations</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-146">3</span><span class="sxs-lookup"><span data-stu-id="3975c-146">3</span></span></p></td>
<td><p><span data-ttu-id="3975c-147">Renvoie les collations de caractères définies dans le catalogue et accessibles par un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-147">Returns the character collations defined in the catalog that are accessible to a given user.</span></span> <span data-ttu-id="3975c-148">(COLLATIONS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-148">(COLLATIONS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-149">COLLATION_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-149">COLLATION_CATALOG</span></span><br />
<span data-ttu-id="3975c-150">COLLATION_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-150">COLLATION_SCHEMA</span></span><br />
<span data-ttu-id="3975c-151">COLLATION_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-151">COLLATION_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-152"><strong>adSchemaColumnPrivileges</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-152"><strong>adSchemaColumnPrivileges</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-153">13</span><span class="sxs-lookup"><span data-stu-id="3975c-153">13</span></span></p></td>
<td><p><span data-ttu-id="3975c-154">Renvoie les privilèges des colonnes des tables définies dans le catalogue,  accordés à un utilisateur donné, ou accordés par ce dernier.
</span><span class="sxs-lookup"><span data-stu-id="3975c-154">Returns the privileges on columns of tables defined in the catalog that are available to, or granted by, a given user.</span></span> <span data-ttu-id="3975c-155">(COLUMN_PRIVILEGES Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-155">(COLUMN_PRIVILEGES Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-156">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-156">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-157">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-157">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-158">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-158">TABLE_NAME</span></span><br />
<span data-ttu-id="3975c-159">NOM_COLONNE</span><span class="sxs-lookup"><span data-stu-id="3975c-159">COLUMN_NAME</span></span><br />
<span data-ttu-id="3975c-160">FOURNISSEUR D’AUTORISATIONS</span><span class="sxs-lookup"><span data-stu-id="3975c-160">GRANTOR</span></span><br />
<span data-ttu-id="3975c-161">BÉNÉFICIAIRE</span><span class="sxs-lookup"><span data-stu-id="3975c-161">GRANTEE</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-162"><strong>adSchemaColumns</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-162"><strong>adSchemaColumns</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-163">4</span><span class="sxs-lookup"><span data-stu-id="3975c-163">4</span></span></p></td>
<td><p><span data-ttu-id="3975c-164">Renvoie les colonnes des tables (vues comprises) définies dans le catalogue et accessibles à un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-164">Returns the columns of tables (including views) defined in the catalog that are accessible to a given user.</span></span> <span data-ttu-id="3975c-165">(COLUMNS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-165">(COLUMNS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-166">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-166">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-167">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-167">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-168">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-168">TABLE_NAME</span></span><br />
<span data-ttu-id="3975c-169">NOM_COLONNE</span><span class="sxs-lookup"><span data-stu-id="3975c-169">COLUMN_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-170"><strong>adSchemaColumnsDomainUsage</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-170"><strong>adSchemaColumnsDomainUsage</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-171">11</span><span class="sxs-lookup"><span data-stu-id="3975c-171">11</span></span></p></td>
<td><p><span data-ttu-id="3975c-172">Renvoie les colonnes définies dans le catalogue et qui dépendent d'un domaine défini dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-172">Returns the columns defined in the catalog that are dependent on a domain defined in the catalog and owned by a given user.</span></span> <span data-ttu-id="3975c-173">(COLUMN_DOMAIN_USAGE Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-173">(COLUMN_DOMAIN_USAGE Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-174">DOMAIN_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-174">DOMAIN_CATALOG</span></span><br />
<span data-ttu-id="3975c-175">DOMAIN_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-175">DOMAIN_SCHEMA</span></span><br />
<span data-ttu-id="3975c-176">NOM_DOMAINE</span><span class="sxs-lookup"><span data-stu-id="3975c-176">DOMAIN_NAME</span></span><br />
<span data-ttu-id="3975c-177">NOM_COLONNE</span><span class="sxs-lookup"><span data-stu-id="3975c-177">COLUMN_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-178"><strong>adSchemaConstraintColumnUsage</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-178"><strong>adSchemaConstraintColumnUsage</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-179">6</span><span class="sxs-lookup"><span data-stu-id="3975c-179">6</span></span></p></td>
<td><p><span data-ttu-id="3975c-180">Renvoie les colonnes utilisées par les contraintes référentielles, uniques et de contrôle, ainsi que par les assertions définies dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-180">Returns the columns used by referential constraints, unique constraints, check constraints, and assertions, defined in the catalog and owned by a given user.</span></span> <span data-ttu-id="3975c-181">(CONSTRAINT_COLUMN_USAGE Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-181">(CONSTRAINT_COLUMN_USAGE Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-182">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-182">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-183">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-183">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-184">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-184">TABLE_NAME</span></span><br />
<span data-ttu-id="3975c-185">NOM_COLONNE</span><span class="sxs-lookup"><span data-stu-id="3975c-185">COLUMN_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-186"><strong>adSchemaConstraintTableUsage</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-186"><strong>adSchemaConstraintTableUsage</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-187">7</span><span class="sxs-lookup"><span data-stu-id="3975c-187">7</span></span></p></td>
<td><p><span data-ttu-id="3975c-188">Renvoie les tables utilisées par les contraintes référentielles, uniques et de contrôle, ainsi que par les assertions définies dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-188">Returns the tables that are used by referential constraints, unique constraints, check constraints, and assertions defined in the catalog and owned by a given user.</span></span> <span data-ttu-id="3975c-189">(CONSTRAINT_TABLE_USAGE Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-189">(CONSTRAINT_TABLE_USAGE Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-190">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-190">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-191">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-191">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-192">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-192">TABLE_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-193"><strong>adSchemaCubes</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-193"><strong>adSchemaCubes</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-194">32</span><span class="sxs-lookup"><span data-stu-id="3975c-194">32</span></span></p></td>
<td><p><span data-ttu-id="3975c-195">Renvoie des informations sur les cubes disponibles dans un schéma (ou catalogue, si le fournisseur ne prend pas en charge les schémas).
</span><span class="sxs-lookup"><span data-stu-id="3975c-195">Returns information about the available cubes in a schema (or the catalog, if the provider does not support schemas).</span></span> <span data-ttu-id="3975c-196">(CUBES Rowset\*)</span><span class="sxs-lookup"><span data-stu-id="3975c-196">(CUBES Rowset\*)</span></span></p></td>
<td><p><span data-ttu-id="3975c-197">CATALOG_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-197">CATALOG_NAME</span></span><br />
<span data-ttu-id="3975c-198">SCHEMA_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-198">SCHEMA_NAME</span></span><br />
<span data-ttu-id="3975c-199">CUBE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-199">CUBE_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-200"><strong>adSchemaDBInfoKeywords</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-200"><strong>adSchemaDBInfoKeywords</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-201">30</span><span class="sxs-lookup"><span data-stu-id="3975c-201">30</span></span></p></td>
<td><p><span data-ttu-id="3975c-202">Renvoie la liste des mots réservés spécifiques aux fournisseur.
</span><span class="sxs-lookup"><span data-stu-id="3975c-202">Returns a list of provider-specific keywords.</span></span> <span data-ttu-id="3975c-203">(IDBInfo::GetKeywords \*)</span><span class="sxs-lookup"><span data-stu-id="3975c-203">(IDBInfo::GetKeywords \*)</span></span></p></td>
<td><p><span data-ttu-id="3975c-204">&lt;Aucun&gt;</span><span class="sxs-lookup"><span data-stu-id="3975c-204">&lt;None&gt;</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-205"><strong>adSchemaDBInfoLiterals</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-205"><strong>adSchemaDBInfoLiterals</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-206">31</span><span class="sxs-lookup"><span data-stu-id="3975c-206">31</span></span></p></td>
<td><p><span data-ttu-id="3975c-207">Renvoie la liste des chaînes littérales spécifiques aux fournisseurs et utilisées dans les commandes texte.
</span><span class="sxs-lookup"><span data-stu-id="3975c-207">Returns a list of provider-specific literals used in text commands.</span></span> <span data-ttu-id="3975c-208">(IDBInfo::GetLiteralInfo \*)</span><span class="sxs-lookup"><span data-stu-id="3975c-208">(IDBInfo::GetLiteralInfo \*)</span></span></p></td>
<td><p><span data-ttu-id="3975c-209">&lt;Aucun&gt;</span><span class="sxs-lookup"><span data-stu-id="3975c-209">&lt;None&gt;</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-210"><strong>adSchemaDimensions</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-210"><strong>adSchemaDimensions</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-211">33</span><span class="sxs-lookup"><span data-stu-id="3975c-211">33</span></span></p></td>
<td><p><span data-ttu-id="3975c-212">Retourne des informations sur les dimensions d’un cube donné.</span><span class="sxs-lookup"><span data-stu-id="3975c-212">Returns information about the dimensions in a given cube.</span></span> <span data-ttu-id="3975c-213">Il comporte une ligne pour chaque dimension.</span><span class="sxs-lookup"><span data-stu-id="3975c-213">It has one row for each dimension.</span></span> <span data-ttu-id="3975c-214">(DIMENSIONS Rowset \*)</span><span class="sxs-lookup"><span data-stu-id="3975c-214">(DIMENSIONS Rowset \*)</span></span></p></td>
<td><p><span data-ttu-id="3975c-215">CATALOG_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-215">CATALOG_NAME</span></span><br />
<span data-ttu-id="3975c-216">SCHEMA_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-216">SCHEMA_NAME</span></span><br />
<span data-ttu-id="3975c-217">CUBE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-217">CUBE_NAME</span></span><br />
<span data-ttu-id="3975c-218">DIMENSION_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-218">DIMENSION_NAME</span></span><br />
<span data-ttu-id="3975c-219">DIMENSION_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-219">DIMENSION_UNIQUE_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-220"><strong>adSchemaForeignKeys</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-220"><strong>adSchemaForeignKeys</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-221">27</span><span class="sxs-lookup"><span data-stu-id="3975c-221">27</span></span></p></td>
<td><p><span data-ttu-id="3975c-222">Renvoie les colonnes de clés étrangères définies dans le catalogue par un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-222">Returns the foreign key columns defined in the catalog by a given user.</span></span> <span data-ttu-id="3975c-223">(FOREIGN_KEYS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-223">(FOREIGN_KEYS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-224">PK_TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-224">PK_TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-225">PK_TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-225">PK_TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-226">PK_TABLE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-226">PK_TABLE_NAME</span></span><br />
<span data-ttu-id="3975c-227">FK_TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-227">FK_TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-228">FK_TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-228">FK_TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-229">FK_TABLE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-229">FK_TABLE_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-230"><strong>adSchemaHierarchies</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-230"><strong>adSchemaHierarchies</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-231">34</span><span class="sxs-lookup"><span data-stu-id="3975c-231">34</span></span></p></td>
<td><p><span data-ttu-id="3975c-232">Renvoie des informations sur les hiérarchies disponibles dans une dimension.
</span><span class="sxs-lookup"><span data-stu-id="3975c-232">Returns information about the hierarchies available in a dimension.</span></span> <span data-ttu-id="3975c-233">(HIERARCHIES Rowset \*)</span><span class="sxs-lookup"><span data-stu-id="3975c-233">(HIERARCHIES Rowset \*)</span></span></p></td>
<td><p><span data-ttu-id="3975c-234">CATALOG_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-234">CATALOG_NAME</span></span><br />
<span data-ttu-id="3975c-235">SCHEMA_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-235">SCHEMA_NAME</span></span><br />
<span data-ttu-id="3975c-236">CUBE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-236">CUBE_NAME</span></span><br />
<span data-ttu-id="3975c-237">DIMENSION_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-237">DIMENSION_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-238">HIERARCHY_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-238">HIERARCHY_NAME</span></span><br />
<span data-ttu-id="3975c-239">HIERARCHY_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-239">HIERARCHY_UNIQUE_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-240"><strong>adSchemaIndexes</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-240"><strong>adSchemaIndexes</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-241">12</span><span class="sxs-lookup"><span data-stu-id="3975c-241">12</span></span></p></td>
<td><p><span data-ttu-id="3975c-242">Renvoie les index définis dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-242">Returns the indexes defined in the catalog that are owned by a given user.</span></span> <span data-ttu-id="3975c-243">(INDEXES Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-243">(INDEXES Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-244">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-244">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-245">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-245">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-246">INDEX_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-246">INDEX_NAME</span></span><br />
<span data-ttu-id="3975c-247">TYPE</span><span class="sxs-lookup"><span data-stu-id="3975c-247">TYPE</span></span><br />
<span data-ttu-id="3975c-248">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-248">TABLE_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-249"><strong>adSchemaKeyColumnUsage</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-249"><strong>adSchemaKeyColumnUsage</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-250">8</span><span class="sxs-lookup"><span data-stu-id="3975c-250">8</span></span></p></td>
<td><p><span data-ttu-id="3975c-251">Renvoie les colonnes définies dans le catalogue et qui sont contraintes sous forme de clés par un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-251">Returns the columns defined in the catalog that are constrained as keys by a given user.</span></span> <span data-ttu-id="3975c-252">(KEY_COLUMN_USAGE Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-252">(KEY_COLUMN_USAGE Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-253">CONSTRAINT_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-253">CONSTRAINT_CATALOG</span></span><br />
<span data-ttu-id="3975c-254">CONSTRAINT_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-254">CONSTRAINT_SCHEMA</span></span><br />
<span data-ttu-id="3975c-255">NOM_CONTRAINTE</span><span class="sxs-lookup"><span data-stu-id="3975c-255">CONSTRAINT_NAME</span></span><br />
<span data-ttu-id="3975c-256">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-256">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-257">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-257">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-258">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-258">TABLE_NAME</span></span><br />
<span data-ttu-id="3975c-259">NOM_COLONNE</span><span class="sxs-lookup"><span data-stu-id="3975c-259">COLUMN_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-260"><strong>adSchemaLevels</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-260"><strong>adSchemaLevels</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-261">35</span><span class="sxs-lookup"><span data-stu-id="3975c-261">35</span></span></p></td>
<td><p><span data-ttu-id="3975c-262">Renvoie des informations sur les niveaux disponibles dans une dimension.
</span><span class="sxs-lookup"><span data-stu-id="3975c-262">Returns information about the levels available in a dimension.</span></span> <span data-ttu-id="3975c-263">(LEVELS Rowset\*)</span><span class="sxs-lookup"><span data-stu-id="3975c-263">(LEVELS Rowset\*)</span></span></p></td>
<td><p><span data-ttu-id="3975c-264">CATALOG_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-264">CATALOG_NAME</span></span><br />
<span data-ttu-id="3975c-265">SCHEMA_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-265">SCHEMA_NAME</span></span><br />
<span data-ttu-id="3975c-266">CUBE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-266">CUBE_NAME</span></span><br />
<span data-ttu-id="3975c-267">DIMENSION_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-267">DIMENSION_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-268">HIERARCHY_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-268">HIERARCHY_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-269">LEVEL_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-269">LEVEL_NAME</span></span><br />
<span data-ttu-id="3975c-270">LEVEL_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-270">LEVEL_UNIQUE_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-271"><strong>adSchemaMeasures</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-271"><strong>adSchemaMeasures</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-272">36</span><span class="sxs-lookup"><span data-stu-id="3975c-272">36</span></span></p></td>
<td><p><span data-ttu-id="3975c-273">Renvoie des informations sur les mesures disponibles.
</span><span class="sxs-lookup"><span data-stu-id="3975c-273">Returns information about the available measures.</span></span> <span data-ttu-id="3975c-274">(MEASURES Rowset \*)</span><span class="sxs-lookup"><span data-stu-id="3975c-274">(MEASURES Rowset \*)</span></span></p></td>
<td><p><span data-ttu-id="3975c-275">CATALOG_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-275">CATALOG_NAME</span></span><br />
<span data-ttu-id="3975c-276">SCHEMA_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-276">SCHEMA_NAME</span></span><br />
<span data-ttu-id="3975c-277">CUBE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-277">CUBE_NAME</span></span><br />
<span data-ttu-id="3975c-278">MEASURE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-278">MEASURE_NAME</span></span><br />
<span data-ttu-id="3975c-279">MEASURE_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-279">MEASURE_UNIQUE_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-280"><strong>adSchemaMembers</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-280"><strong>adSchemaMembers</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-281">38</span><span class="sxs-lookup"><span data-stu-id="3975c-281">38</span></span></p></td>
<td><p><span data-ttu-id="3975c-282">Renvoie des informations sur les membres disponibles.
</span><span class="sxs-lookup"><span data-stu-id="3975c-282">Returns information about the available members.</span></span> <span data-ttu-id="3975c-283">(MEMBERS Rowset \*)</span><span class="sxs-lookup"><span data-stu-id="3975c-283">(MEMBERS Rowset \*)</span></span></p></td>
<td><p><span data-ttu-id="3975c-284">CATALOG_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-284">CATALOG_NAME</span></span><br />
<span data-ttu-id="3975c-285">SCHEMA_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-285">SCHEMA_NAME</span></span><br />
<span data-ttu-id="3975c-286">CUBE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-286">CUBE_NAME</span></span><br />
<span data-ttu-id="3975c-287">DIMENSION_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-287">DIMENSION_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-288">HIERARCHY_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-288">HIERARCHY_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-289">LEVEL_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-289">LEVEL_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-290">LEVEL_NUMBER</span><span class="sxs-lookup"><span data-stu-id="3975c-290">LEVEL_NUMBER</span></span><br />
<span data-ttu-id="3975c-291">MEMBER_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-291">MEMBER_NAME</span></span><br />
<span data-ttu-id="3975c-292">MEMBER_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-292">MEMBER_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-293">MEMBER_CAPTION</span><span class="sxs-lookup"><span data-stu-id="3975c-293">MEMBER_CAPTION</span></span><br />
<span data-ttu-id="3975c-294">ARBRE</span><span class="sxs-lookup"><span data-stu-id="3975c-294">MEMBER_TYPE</span></span><br />
<span data-ttu-id="3975c-295">Opérateur d’arbre (pour plus d’informations, consultez OLE DB pour la documentation OLAP.)</span><span class="sxs-lookup"><span data-stu-id="3975c-295">Tree operator (For more information, see the OLE DB for OLAP documentation.)</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-296"><strong>adSchemaPrimaryKeys</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-296"><strong>adSchemaPrimaryKeys</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-297">28</span><span class="sxs-lookup"><span data-stu-id="3975c-297">28</span></span></p></td>
<td><p><span data-ttu-id="3975c-298">Renvoie les colonnes de clés primaires définies dans le catalogue par un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-298">Returns the primary key columns defined in the catalog by a given user.</span></span> <span data-ttu-id="3975c-299">(PRIMARY_KEYS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-299">(PRIMARY_KEYS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-300">PK_TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-300">PK_TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-301">PK_TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-301">PK_TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-302">PK_TABLE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-302">PK_TABLE_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-303"><strong>adSchemaProcedureColumns</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-303"><strong>adSchemaProcedureColumns</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-304">29</span><span class="sxs-lookup"><span data-stu-id="3975c-304">29</span></span></p></td>
<td><p><span data-ttu-id="3975c-305">Renvoie des informations sur les colonnes de jeux de lignes renvoyées par des procédures.
</span><span class="sxs-lookup"><span data-stu-id="3975c-305">Returns information about the columns of rowsets returned by procedures.</span></span> <span data-ttu-id="3975c-306">(PROCEDURE_COLUMNS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-306">(PROCEDURE_COLUMNS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-307">PROCEDURE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-307">PROCEDURE_CATALOG</span></span><br />
<span data-ttu-id="3975c-308">PROCEDURE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-308">PROCEDURE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-309">NOM_DE_PROCÉDURE</span><span class="sxs-lookup"><span data-stu-id="3975c-309">PROCEDURE_NAME</span></span><br />
<span data-ttu-id="3975c-310">NOM_COLONNE</span><span class="sxs-lookup"><span data-stu-id="3975c-310">COLUMN_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-311"><strong>adSchemaProcedureParameters</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-311"><strong>adSchemaProcedureParameters</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-312">26</span><span class="sxs-lookup"><span data-stu-id="3975c-312">26</span></span></p></td>
<td><p><span data-ttu-id="3975c-313">Renvoie des informations sur les paramètres et les codes de retour des procédures.
</span><span class="sxs-lookup"><span data-stu-id="3975c-313">Returns information about the parameters and return codes of procedures.</span></span> <span data-ttu-id="3975c-314">(PROCEDURE_PARAMETERS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-314">(PROCEDURE_PARAMETERS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-315">PROCEDURE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-315">PROCEDURE_CATALOG</span></span><br />
<span data-ttu-id="3975c-316">PROCEDURE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-316">PROCEDURE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-317">NOM_DE_PROCÉDURE</span><span class="sxs-lookup"><span data-stu-id="3975c-317">PROCEDURE_NAME</span></span><br />
<span data-ttu-id="3975c-318">PARAMETER_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-318">PARAMETER_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-319"><strong>adSchemaProcedures</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-319"><strong>adSchemaProcedures</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-320">16</span><span class="sxs-lookup"><span data-stu-id="3975c-320">16</span></span></p></td>
<td><p><span data-ttu-id="3975c-321">Renvoie les procédures définies dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-321">Returns the procedures defined in the catalog that are owned by a given user.</span></span> <span data-ttu-id="3975c-322">(PROCEDURES Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-322">(PROCEDURES Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-323">PROCEDURE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-323">PROCEDURE_CATALOG</span></span><br />
<span data-ttu-id="3975c-324">PROCEDURE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-324">PROCEDURE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-325">NOM_DE_PROCÉDURE</span><span class="sxs-lookup"><span data-stu-id="3975c-325">PROCEDURE_NAME</span></span><br />
<span data-ttu-id="3975c-326">PROCEDURE_TYPE</span><span class="sxs-lookup"><span data-stu-id="3975c-326">PROCEDURE_TYPE</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-327"><strong>adSchemaProperties</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-327"><strong>adSchemaProperties</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-328">37</span><span class="sxs-lookup"><span data-stu-id="3975c-328">37</span></span></p></td>
<td><p><span data-ttu-id="3975c-329">Renvoie des informations sur les propriétés disponibles pour chaque niveau de la dimension.
</span><span class="sxs-lookup"><span data-stu-id="3975c-329">Returns information about the available properties for each level of the dimension.</span></span> <span data-ttu-id="3975c-330">(PROPERTIES Rowset \*)</span><span class="sxs-lookup"><span data-stu-id="3975c-330">(PROPERTIES Rowset \*)</span></span></p></td>
<td><p><span data-ttu-id="3975c-331">CATALOG_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-331">CATALOG_NAME</span></span><br />
<span data-ttu-id="3975c-332">SCHEMA_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-332">SCHEMA_NAME</span></span><br />
<span data-ttu-id="3975c-333">CUBE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-333">CUBE_NAME</span></span><br />
<span data-ttu-id="3975c-334">DIMENSION_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-334">DIMENSION_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-335">HIERARCHY_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-335">HIERARCHY_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-336">LEVEL_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-336">LEVEL_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-337">MEMBER_UNIQUE_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-337">MEMBER_UNIQUE_NAME</span></span><br />
<span data-ttu-id="3975c-338">PROPERTY_TYPE</span><span class="sxs-lookup"><span data-stu-id="3975c-338">PROPERTY_TYPE</span></span><br />
<span data-ttu-id="3975c-339">PROPERTY_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-339">PROPERTY_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-340"><strong>adSchemaProviderSpecific</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-340"><strong>adSchemaProviderSpecific</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-341">-1</span><span class="sxs-lookup"><span data-stu-id="3975c-341">-1</span></span></p></td>
<td><p><span data-ttu-id="3975c-342">Utilisé si le fournisseur définit ses propres requêtes de schéma non standard.</span><span class="sxs-lookup"><span data-stu-id="3975c-342">Used if the provider defines its own nonstandard schema queries.</span></span></p></td>
<td><p><span data-ttu-id="3975c-343">&lt;Spécifique au fournisseur&gt;</span><span class="sxs-lookup"><span data-stu-id="3975c-343">&lt;Provider specific&gt;</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-344"><strong>adSchemaProviderTypes</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-344"><strong>adSchemaProviderTypes</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-345">22</span><span class="sxs-lookup"><span data-stu-id="3975c-345">22</span></span></p></td>
<td><p><span data-ttu-id="3975c-346">Renvoie les types de données (primaires) pris en charge par le fournisseur de données.
</span><span class="sxs-lookup"><span data-stu-id="3975c-346">Returns the (base) data types supported by the data provider.</span></span> <span data-ttu-id="3975c-347">(PROVIDER_TYPES Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-347">(PROVIDER_TYPES Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-348">TYPE_DE_DONNÉES</span><span class="sxs-lookup"><span data-stu-id="3975c-348">DATA_TYPE</span></span><br />
<span data-ttu-id="3975c-349">BEST_MATCH</span><span class="sxs-lookup"><span data-stu-id="3975c-349">BEST_MATCH</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-350"><strong>AdSchemaReferentialConstraints</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-350"><strong>AdSchemaReferentialConstraints</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-351">9</span><span class="sxs-lookup"><span data-stu-id="3975c-351">9</span></span></p></td>
<td><p><span data-ttu-id="3975c-352">Renvoie les contraintes référentielles définies dans le catalogie et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-352">Returns the referential constraints defined in the catalog that are owned by a given user.</span></span> <span data-ttu-id="3975c-353">(REFERENTIAL_CONSTRAINTS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-353">(REFERENTIAL_CONSTRAINTS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-354">CONSTRAINT_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-354">CONSTRAINT_CATALOG</span></span><br />
<span data-ttu-id="3975c-355">CONSTRAINT_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-355">CONSTRAINT_SCHEMA</span></span><br />
<span data-ttu-id="3975c-356">NOM_CONTRAINTE</span><span class="sxs-lookup"><span data-stu-id="3975c-356">CONSTRAINT_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-357"><strong>adSchemaSchemata</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-357"><strong>adSchemaSchemata</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-358">17</span><span class="sxs-lookup"><span data-stu-id="3975c-358">17</span></span></p></td>
<td><p><span data-ttu-id="3975c-359">Renvoie les schémas (objets de la base de données) dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-359">Returns the schemas (database objects) that are owned by a given user.</span></span> <span data-ttu-id="3975c-360">(SCHEMATA Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-360">(SCHEMATA Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-361">CATALOG_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-361">CATALOG_NAME</span></span><br />
<span data-ttu-id="3975c-362">SCHEMA_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-362">SCHEMA_NAME</span></span><br />
<span data-ttu-id="3975c-363">SCHEMA_OWNER</span><span class="sxs-lookup"><span data-stu-id="3975c-363">SCHEMA_OWNER</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-364"><strong>adSchemaSQLLanguages</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-364"><strong>adSchemaSQLLanguages</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-365">18</span><span class="sxs-lookup"><span data-stu-id="3975c-365">18</span></span></p></td>
<td><p><span data-ttu-id="3975c-366">Renvoie les niveaux de conformité, les options et les dialectes pris en charge par les données de traitement d'implémentation SQL définies dans le catalogue.
</span><span class="sxs-lookup"><span data-stu-id="3975c-366">Returns the conformance levels, options, and dialects supported by the SQL-implementation processing data defined in the catalog.</span></span> <span data-ttu-id="3975c-367">(SQL_LANGUAGES Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-367">(SQL_LANGUAGES Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-368">&lt;Aucun&gt;</span><span class="sxs-lookup"><span data-stu-id="3975c-368">&lt;None&gt;</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-369"><strong>adSchemaStatistics</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-369"><strong>adSchemaStatistics</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-370">19</span><span class="sxs-lookup"><span data-stu-id="3975c-370">19</span></span></p></td>
<td><p><span data-ttu-id="3975c-371">Renvoie les statistiques définies dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-371">Returns the statistics defined in the catalog that are owned by a given user.</span></span> <span data-ttu-id="3975c-372">(STATISTICS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-372">(STATISTICS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-373">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-373">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-374">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-374">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-375">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-375">TABLE_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-376"><strong>adSchemaTableConstraints</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-376"><strong>adSchemaTableConstraints</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-377">10</span><span class="sxs-lookup"><span data-stu-id="3975c-377">10</span></span></p></td>
<td><p><span data-ttu-id="3975c-378">Renvoie les contraintes de table définies dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-378">Returns the table constraints defined in the catalog that are owned by a given user.</span></span> <span data-ttu-id="3975c-379">(TABLE_CONSTRAINTS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-379">(TABLE_CONSTRAINTS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-380">CONSTRAINT_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-380">CONSTRAINT_CATALOG</span></span><br />
<span data-ttu-id="3975c-381">CONSTRAINT_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-381">CONSTRAINT_SCHEMA</span></span><br />
<span data-ttu-id="3975c-382">NOM_CONTRAINTE</span><span class="sxs-lookup"><span data-stu-id="3975c-382">CONSTRAINT_NAME</span></span><br />
<span data-ttu-id="3975c-383">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-383">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-384">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-384">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-385">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-385">TABLE_NAME</span></span><br />
<span data-ttu-id="3975c-386">CONSTRAINT_TYPE</span><span class="sxs-lookup"><span data-stu-id="3975c-386">CONSTRAINT_TYPE</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-387"><strong>adSchemaTablePrivileges</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-387"><strong>adSchemaTablePrivileges</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-388">14</span><span class="sxs-lookup"><span data-stu-id="3975c-388">14</span></span></p></td>
<td><p><span data-ttu-id="3975c-389">Renvoie les privilèges sur les tables définies dans le catalogue et qui sont disponibles pour un utilisateur donné, ou accordés par ce dernier.
</span><span class="sxs-lookup"><span data-stu-id="3975c-389">Returns the privileges on tables defined in the catalog that are available to, or granted by, a given user.</span></span> <span data-ttu-id="3975c-390">(TABLE_PRIVILEGES Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-390">(TABLE_PRIVILEGES Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-391">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-391">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-392">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-392">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-393">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-393">TABLE_NAME</span></span><br />
<span data-ttu-id="3975c-394">FOURNISSEUR D’AUTORISATIONS</span><span class="sxs-lookup"><span data-stu-id="3975c-394">GRANTOR</span></span><br />
<span data-ttu-id="3975c-395">BÉNÉFICIAIRE</span><span class="sxs-lookup"><span data-stu-id="3975c-395">GRANTEE</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-396"><strong>requêtes adSchemaTables</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-396"><strong>adSchemaTables</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-397">20</span><span class="sxs-lookup"><span data-stu-id="3975c-397">20</span></span></p></td>
<td><p><span data-ttu-id="3975c-398">Renvoie les tables (notamment les vues) définies dans le catalogue et accessibles par un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-398">Returns the tables (including views) defined in the catalog that are accessible to a given user.</span></span> <span data-ttu-id="3975c-399">(TABLES Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-399">(TABLES Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-400">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-400">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-401">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-401">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-402">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-402">TABLE_NAME</span></span><br />
<span data-ttu-id="3975c-403">TABLE_TYPE</span><span class="sxs-lookup"><span data-stu-id="3975c-403">TABLE_TYPE</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-404"><strong>adSchemaTranslations</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-404"><strong>adSchemaTranslations</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-405">21</span><span class="sxs-lookup"><span data-stu-id="3975c-405">21</span></span></p></td>
<td><p><span data-ttu-id="3975c-406">Renvoie les conversions de caractères définies dans le catalogue et auxquelles un utilisateur donné a accès.
</span><span class="sxs-lookup"><span data-stu-id="3975c-406">Returns the character translations defined in the catalog that are accessible to a given user.</span></span> <span data-ttu-id="3975c-407">(TRANSLATIONS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-407">(TRANSLATIONS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-408">TRANSLATION_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-408">TRANSLATION_CATALOG</span></span><br />
<span data-ttu-id="3975c-409">TRANSLATION_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-409">TRANSLATION_SCHEMA</span></span><br />
<span data-ttu-id="3975c-410">TRANSLATION_NAME</span><span class="sxs-lookup"><span data-stu-id="3975c-410">TRANSLATION_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-411"><strong>adSchemaTrustees</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-411"><strong>adSchemaTrustees</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-412">39</span><span class="sxs-lookup"><span data-stu-id="3975c-412">39</span></span></p></td>
<td><p><span data-ttu-id="3975c-413">Réservé pour usage ultérieur.</span><span class="sxs-lookup"><span data-stu-id="3975c-413">Reserved for future use.</span></span></p></td>
<td><p><br />
</p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-414"><strong>adSchemaUsagePrivileges</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-414"><strong>adSchemaUsagePrivileges</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-415">15</span><span class="sxs-lookup"><span data-stu-id="3975c-415">15</span></span></p></td>
<td><p><span data-ttu-id="3975c-416">Renvoie les privilèges USAGE sur les objets définis dans le catalogue et qui sont disponibles pour un utilisateur donné, ou accordés par ce dernier.
</span><span class="sxs-lookup"><span data-stu-id="3975c-416">Returns the USAGE privileges on objects defined in the catalog that are available to, or granted by, a given user.</span></span> <span data-ttu-id="3975c-417">(USAGE_PRIVILEGES Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-417">(USAGE_PRIVILEGES Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-418">OBJECT_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-418">OBJECT_CATALOG</span></span><br />
<span data-ttu-id="3975c-419">OBJECT_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-419">OBJECT_SCHEMA</span></span><br />
<span data-ttu-id="3975c-420">NOM_DE_L</span><span class="sxs-lookup"><span data-stu-id="3975c-420">OBJECT_NAME</span></span><br />
<span data-ttu-id="3975c-421">OBJECT_TYPE</span><span class="sxs-lookup"><span data-stu-id="3975c-421">OBJECT_TYPE</span></span><br />
<span data-ttu-id="3975c-422">FOURNISSEUR D’AUTORISATIONS</span><span class="sxs-lookup"><span data-stu-id="3975c-422">GRANTOR</span></span><br />
<span data-ttu-id="3975c-423">BÉNÉFICIAIRE</span><span class="sxs-lookup"><span data-stu-id="3975c-423">GRANTEE</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-424"><strong>adSchemaViewColumnUsage</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-424"><strong>adSchemaViewColumnUsage</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-425">24</span><span class="sxs-lookup"><span data-stu-id="3975c-425">24</span></span></p></td>
<td><p><span data-ttu-id="3975c-426">Renvoie les colonnes desquelles dépendent les tables vues, définies dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-426">Returns the columns on which viewed tables, defined in the catalog and owned by a given user, are dependent.</span></span> <span data-ttu-id="3975c-427">(VIEW_COLUMN_USAGE Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-427">(VIEW_COLUMN_USAGE Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-428">VIEW_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-428">VIEW_CATALOG</span></span><br />
<span data-ttu-id="3975c-429">VIEW_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-429">VIEW_SCHEMA</span></span><br />
<span data-ttu-id="3975c-430">NOM_DE_VUE</span><span class="sxs-lookup"><span data-stu-id="3975c-430">VIEW_NAME</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-431"><strong>adSchemaViews</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-431"><strong>adSchemaViews</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-432">23</span><span class="sxs-lookup"><span data-stu-id="3975c-432">23</span></span></p></td>
<td><p><span data-ttu-id="3975c-433">Renvoie les vues définies dans le catalogue et accessibles par un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-433">Returns the views defined in the catalog that are accessible to a given user.</span></span> <span data-ttu-id="3975c-434">(VIEWS Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-434">(VIEWS Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-435">TABLE_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-435">TABLE_CATALOG</span></span><br />
<span data-ttu-id="3975c-436">TABLE_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-436">TABLE_SCHEMA</span></span><br />
<span data-ttu-id="3975c-437">NOM_DE_TABLE</span><span class="sxs-lookup"><span data-stu-id="3975c-437">TABLE_NAME</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-438"><strong>adSchemaViewTableUsage</strong></span><span class="sxs-lookup"><span data-stu-id="3975c-438"><strong>adSchemaViewTableUsage</strong></span></span></p></td>
<td><p><span data-ttu-id="3975c-439">25</span><span class="sxs-lookup"><span data-stu-id="3975c-439">25</span></span></p></td>
<td><p><span data-ttu-id="3975c-440">Renvoie les tables desquelles dépendent les tables vues, définies dans le catalogue et dont est propriétaire un utilisateur donné.
</span><span class="sxs-lookup"><span data-stu-id="3975c-440">Returns the tables on which viewed tables, defined in the catalog and owned by a given user, are dependent.</span></span> <span data-ttu-id="3975c-441">(VIEW_TABLE_USAGE Rowset)</span><span class="sxs-lookup"><span data-stu-id="3975c-441">(VIEW_TABLE_USAGE Rowset)</span></span></p></td>
<td><p><span data-ttu-id="3975c-442">VIEW_CATALOG</span><span class="sxs-lookup"><span data-stu-id="3975c-442">VIEW_CATALOG</span></span><br />
<span data-ttu-id="3975c-443">VIEW_SCHEMA</span><span class="sxs-lookup"><span data-stu-id="3975c-443">VIEW_SCHEMA</span></span><br />
<span data-ttu-id="3975c-444">NOM_DE_VUE</span><span class="sxs-lookup"><span data-stu-id="3975c-444">VIEW_NAME</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="3975c-445">**Équivalent ADO/WFC**</span><span class="sxs-lookup"><span data-stu-id="3975c-445">**ADO/WFC Equivalent**</span></span>

<span data-ttu-id="3975c-446">Module : **com.ms.wfc.data**</span><span class="sxs-lookup"><span data-stu-id="3975c-446">Package: **com.ms.wfc.data**</span></span>

<table>
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="3975c-447">Constante</span><span class="sxs-lookup"><span data-stu-id="3975c-447">Constant</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="3975c-448">AdoEnums.Schema.ASSERTS</span><span class="sxs-lookup"><span data-stu-id="3975c-448">AdoEnums.Schema.ASSERTS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-449">AdoEnums.Schema.CATALOGS</span><span class="sxs-lookup"><span data-stu-id="3975c-449">AdoEnums.Schema.CATALOGS</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-450">AdoEnums.Schema.CHARACTERSETS</span><span class="sxs-lookup"><span data-stu-id="3975c-450">AdoEnums.Schema.CHARACTERSETS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-451">AdoEnums.Schema.CHECKCONSTRAINTS</span><span class="sxs-lookup"><span data-stu-id="3975c-451">AdoEnums.Schema.CHECKCONSTRAINTS</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-452">AdoEnums.Schema.COLLATIONS</span><span class="sxs-lookup"><span data-stu-id="3975c-452">AdoEnums.Schema.COLLATIONS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-453">AdoEnums.Schema.COLUMNPRIVILEGES</span><span class="sxs-lookup"><span data-stu-id="3975c-453">AdoEnums.Schema.COLUMNPRIVILEGES</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-454">AdoEnums.Schema.COLUMNS</span><span class="sxs-lookup"><span data-stu-id="3975c-454">AdoEnums.Schema.COLUMNS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-455">AdoEnums.Schema.COLUMNSDOMAINUSAGE</span><span class="sxs-lookup"><span data-stu-id="3975c-455">AdoEnums.Schema.COLUMNSDOMAINUSAGE</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-456">AdoEnums.Schema.CONSTRAINTCOLUMNUSAGE</span><span class="sxs-lookup"><span data-stu-id="3975c-456">AdoEnums.Schema.CONSTRAINTCOLUMNUSAGE</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-457">AdoEnums.Schema.CONSTRAINTTABLEUSAGE</span><span class="sxs-lookup"><span data-stu-id="3975c-457">AdoEnums.Schema.CONSTRAINTTABLEUSAGE</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-458">AdoEnums.Schema.CUBES</span><span class="sxs-lookup"><span data-stu-id="3975c-458">AdoEnums.Schema.CUBES</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-459">AdoEnums.Schema.DBINFOKEYWORDS</span><span class="sxs-lookup"><span data-stu-id="3975c-459">AdoEnums.Schema.DBINFOKEYWORDS</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-460">AdoEnums.Schema.DBINFOLITERALS</span><span class="sxs-lookup"><span data-stu-id="3975c-460">AdoEnums.Schema.DBINFOLITERALS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-461">AdoEnums.Schema.DIMENSIONS</span><span class="sxs-lookup"><span data-stu-id="3975c-461">AdoEnums.Schema.DIMENSIONS</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-462">AdoEnums.Schema.FOREIGNKEYS</span><span class="sxs-lookup"><span data-stu-id="3975c-462">AdoEnums.Schema.FOREIGNKEYS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-463">AdoEnums.Schema.HIERARCHIES</span><span class="sxs-lookup"><span data-stu-id="3975c-463">AdoEnums.Schema.HIERARCHIES</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-464">AdoEnums.Schema.INDEXES</span><span class="sxs-lookup"><span data-stu-id="3975c-464">AdoEnums.Schema.INDEXES</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-465">AdoEnums.Schema.KEYCOLUMNUSAGE</span><span class="sxs-lookup"><span data-stu-id="3975c-465">AdoEnums.Schema.KEYCOLUMNUSAGE</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-466">AdoEnums.Schema.LEVELS</span><span class="sxs-lookup"><span data-stu-id="3975c-466">AdoEnums.Schema.LEVELS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-467">AdoEnums.Schema.MEASURES</span><span class="sxs-lookup"><span data-stu-id="3975c-467">AdoEnums.Schema.MEASURES</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-468">AdoEnums.Schema.MEMBERS</span><span class="sxs-lookup"><span data-stu-id="3975c-468">AdoEnums.Schema.MEMBERS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-469">AdoEnums.Schema.PRIMARYKEYS</span><span class="sxs-lookup"><span data-stu-id="3975c-469">AdoEnums.Schema.PRIMARYKEYS</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-470">AdoEnums.Schema.PROCEDURECOLUMNS</span><span class="sxs-lookup"><span data-stu-id="3975c-470">AdoEnums.Schema.PROCEDURECOLUMNS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-471">AdoEnums.Schema.PROCEDUREPARAMETERS</span><span class="sxs-lookup"><span data-stu-id="3975c-471">AdoEnums.Schema.PROCEDUREPARAMETERS</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-472">AdoEnums.Schema.PROCEDURES</span><span class="sxs-lookup"><span data-stu-id="3975c-472">AdoEnums.Schema.PROCEDURES</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-473">AdoEnums.Schema.PROPERTIES</span><span class="sxs-lookup"><span data-stu-id="3975c-473">AdoEnums.Schema.PROPERTIES</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-474">AdoEnums.Schema.PROVIDERSPECIFIC</span><span class="sxs-lookup"><span data-stu-id="3975c-474">AdoEnums.Schema.PROVIDERSPECIFIC</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-475">AdoEnums.Schema.PROVIDERTYPES</span><span class="sxs-lookup"><span data-stu-id="3975c-475">AdoEnums.Schema.PROVIDERTYPES</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-476">AdoEnums.Schema.REFERENTIALCONTRAINTS</span><span class="sxs-lookup"><span data-stu-id="3975c-476">AdoEnums.Schema.REFERENTIALCONTRAINTS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-477">AdoEnums.Schema.SCHEMATA</span><span class="sxs-lookup"><span data-stu-id="3975c-477">AdoEnums.Schema.SCHEMATA</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-478">AdoEnums.Schema.SQLLANGUAGES</span><span class="sxs-lookup"><span data-stu-id="3975c-478">AdoEnums.Schema.SQLLANGUAGES</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-479">AdoEnums.Schema.STATISTICS</span><span class="sxs-lookup"><span data-stu-id="3975c-479">AdoEnums.Schema.STATISTICS</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-480">AdoEnums.Schema.TABLECONSTRAINTS</span><span class="sxs-lookup"><span data-stu-id="3975c-480">AdoEnums.Schema.TABLECONSTRAINTS</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-481">AdoEnums.Schema.TABLEPRIVILEGES</span><span class="sxs-lookup"><span data-stu-id="3975c-481">AdoEnums.Schema.TABLEPRIVILEGES</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-482">AdoEnums.Schema.TABLES</span><span class="sxs-lookup"><span data-stu-id="3975c-482">AdoEnums.Schema.TABLES</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-483">AdoEnums.Schema.TRANSLATIONS</span><span class="sxs-lookup"><span data-stu-id="3975c-483">AdoEnums.Schema.TRANSLATIONS</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-484">AdoEnums.Schema.TRUSTEES</span><span class="sxs-lookup"><span data-stu-id="3975c-484">AdoEnums.Schema.TRUSTEES</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-485">AdoEnums.Schema.USAGEPRIVILEGES</span><span class="sxs-lookup"><span data-stu-id="3975c-485">AdoEnums.Schema.USAGEPRIVILEGES</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-486">AdoEnums.Schema.VIEWCOLUMNUSAGE</span><span class="sxs-lookup"><span data-stu-id="3975c-486">AdoEnums.Schema.VIEWCOLUMNUSAGE</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="3975c-487">AdoEnums.Schema.VIEWS</span><span class="sxs-lookup"><span data-stu-id="3975c-487">AdoEnums.Schema.VIEWS</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="3975c-488">AdoEnums.Schema.VIEWTABLEUSAGE</span><span class="sxs-lookup"><span data-stu-id="3975c-488">AdoEnums.Schema.VIEWTABLEUSAGE</span></span></p></td>
</tr>
</tbody>
</table>
