---
title: DBEngine.CreateDatabase Method (DAO)
TOCTitle: CreateDatabase Method
ms:assetid: d5821a4b-483a-b8fa-e929-5f036057d8c4
ms:mtpsurl: https://msdn.microsoft.com/library/Ff835033(v=office.15)
ms:contentKeyID: 48547966
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- dao360.chm1052972
f1_categories:
- Office.Version=v15
ms.openlocfilehash: f1f47e403b1b1c61bb766d28be05558f28dba8f1
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470098"
---
# <a name="dbenginecreatedatabase-method-dao"></a><span data-ttu-id="1b071-102">DBEngine.CreateDatabase Method (DAO)</span><span class="sxs-lookup"><span data-stu-id="1b071-102">DBEngine.CreateDatabase Method (DAO)</span></span>


<span data-ttu-id="1b071-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="1b071-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="1b071-p101">Crée un objet **[Database](database-object-dao.md)**, enregistre la base de données sur le disque, et renvoie un objet **Database** ouvert (espaces de travail Microsoft Access uniquement).</span><span class="sxs-lookup"><span data-stu-id="1b071-p101">Creates a new **[Database](database-object-dao.md)** object, saves the database to disk, and returns an opened **Database** object (Microsoft Access workspaces only). .</span></span>

## <a name="syntax"></a><span data-ttu-id="1b071-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="1b071-106">Syntax</span></span>

<span data-ttu-id="1b071-107">*expression* . CreateDatabase (***nom***, ***paramètres régionaux***, ***Option***)</span><span class="sxs-lookup"><span data-stu-id="1b071-107">*expression* .CreateDatabase(***Name***, ***Locale***, ***Option***)</span></span>

<span data-ttu-id="1b071-108">*expression* Variable qui représente un objet **DBEngine** .</span><span class="sxs-lookup"><span data-stu-id="1b071-108">*expression* A variable that represents a **DBEngine** object.</span></span>

### <a name="parameters"></a><span data-ttu-id="1b071-109">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1b071-109">Parameters</span></span>

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="1b071-110">Name</span><span class="sxs-lookup"><span data-stu-id="1b071-110">Name</span></span></p></th>
<th><p><span data-ttu-id="1b071-111">Obligatoire/Facultatif</span><span class="sxs-lookup"><span data-stu-id="1b071-111">Required/Optional</span></span></p></th>
<th><p><span data-ttu-id="1b071-112">Type de données</span><span class="sxs-lookup"><span data-stu-id="1b071-112">Data Type</span></span></p></th>
<th><p><span data-ttu-id="1b071-113">Description</span><span class="sxs-lookup"><span data-stu-id="1b071-113">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="1b071-114">Name</span><span class="sxs-lookup"><span data-stu-id="1b071-114">Name</span></span></p></td>
<td><p><span data-ttu-id="1b071-115">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="1b071-115">Required</span></span></p></td>
<td><p><span data-ttu-id="1b071-116"><strong>Chaîne</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-116"><strong>String</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-117">Une chaîne ne pouvant dépasser 255 caractères qui est le nom du fichier de base de données que vous créez.</span><span class="sxs-lookup"><span data-stu-id="1b071-117">A String up to 255 characters long that is the name of the database file that you're creating.</span></span> <span data-ttu-id="1b071-118">Il peut être le nom de fichier et chemin d’accès complet.</span><span class="sxs-lookup"><span data-stu-id="1b071-118">It can be the full path and file name.</span></span> <span data-ttu-id="1b071-119">Si votre réseau prend en charge, vous pouvez également spécifier un chemin d’accès réseau, tel que &quot; \\server1\share1\dir1\db1&quot;.</span><span class="sxs-lookup"><span data-stu-id="1b071-119">If your network supports it, you can also specify a network path, such as &quot;\\server1\share1\dir1\db1&quot;.</span></span> <span data-ttu-id="1b071-120">Vous pouvez uniquement créer des fichiers de base de données Microsoft Access avec cette méthode.</span><span class="sxs-lookup"><span data-stu-id="1b071-120">You can only create Microsoft Access database files with this method.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-121">Locale</span><span class="sxs-lookup"><span data-stu-id="1b071-121">Locale</span></span></p></td>
<td><p><span data-ttu-id="1b071-122">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="1b071-122">Required</span></span></p></td>
<td><p><span data-ttu-id="1b071-123"><strong>Chaîne</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-123"><strong>String</strong></span></span></p></td>
<td><ul>
<li><p><span data-ttu-id="1b071-p103">Expression de chaîne qui spécifie un ordre de classement pour la création de la base de données, tel qu'il est spécifié dans la section Remarques. Vous devez indiquer cet argument sans quoi une erreur se produit.</span><span class="sxs-lookup"><span data-stu-id="1b071-p103">A string expression that specifies a collating order for creating the database, as specified in Settings. You must supply this argument or an error occurs.</span></span></p></li>
<li><p><span data-ttu-id="1b071-126">Vous pouvez également créer un mot de passe pour le nouvel objet <strong>Database</strong> en concaténant la chaîne de mot de passe (commençant par &quot;; pwd =&quot; ) avec une constante dans l’argument <em>locale</em> , comme suit :</span><span class="sxs-lookup"><span data-stu-id="1b071-126">You can also create a password for the new <strong>Database</strong> object by concatenating the password string (starting with &quot;;pwd=&quot; ) with a constant in the <em>locale</em> argument, like this:</span></span></p></li>
<li><p><span data-ttu-id="1b071-127">dbLangSpanish &amp; &quot;; pwd = NewPassword&quot;</span><span class="sxs-lookup"><span data-stu-id="1b071-127">dbLangSpanish &amp; &quot;;pwd=NewPassword&quot;</span></span></p></li>
<li><p><span data-ttu-id="1b071-128">Si vous souhaitez utiliser la valeur par défaut de <em>locale</em> mais spécifier un mot de passe, entrez simplement une chaîne de mot de passe pour l'argument <em>locale</em> :</span><span class="sxs-lookup"><span data-stu-id="1b071-128">If you want to use the default <em>locale</em>, but specify a password, simply enter a password string for the <em>locale</em> argument:</span></span></p></li>
<li><p><span data-ttu-id="1b071-129">&quot;; pwd = NewPassword&quot;</span><span class="sxs-lookup"><span data-stu-id="1b071-129">&quot;;pwd=NewPassword&quot;</span></span></p></li>
<li><p><span data-ttu-id="1b071-p104">[!REMARQUE] Définissez des mots de passe forts qui combinent des lettres minuscules et majuscules, des nombres et des symboles. Les mots de passe faibles ne regroupent pas ces éléments. Mot de passe fort : Y6dh!et5. Mot de passe faible : Maison27. Utilisez un mot de passe fort dont vous pouvez vous souvenir sans devoir le noter.</span><span class="sxs-lookup"><span data-stu-id="1b071-p104">Use strong passwords that combine upper- and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27. Use a strong password that you can remember so that you don't have to write it down.</span></span></p></li>
</ul></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-135">Option</span><span class="sxs-lookup"><span data-stu-id="1b071-135">Option</span></span></p></td>
<td><p><span data-ttu-id="1b071-136">Facultatif</span><span class="sxs-lookup"><span data-stu-id="1b071-136">Optional</span></span></p></td>
<td><p><span data-ttu-id="1b071-137"><strong>Variante</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-137"><strong>Variant</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-p105">Constante ou combinaison de constantes qui indique une ou plusieurs options, comme spécifié dans la section Remarques. Vous pouvez combiner des options en associant les constantes correspondantes.</span><span class="sxs-lookup"><span data-stu-id="1b071-p105">A constant or combination of constants that indicates one or more options, as specified in Settings. You can combine options by summing the corresponding constants.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="1b071-140">Remarques</span><span class="sxs-lookup"><span data-stu-id="1b071-140">Remarks</span></span>

<span data-ttu-id="1b071-141">Vous pouvez utiliser l’une des constantes suivantes pour l’argument ParamètresRégionaux afin de spécifier la propriété **[CollatingOrder](database-collatingorder-property-dao.md)** du texte pour des comparaisons de chaînes.</span><span class="sxs-lookup"><span data-stu-id="1b071-141">You can use one of the following constants for the locale argument to specify the **[CollatingOrder](database-collatingorder-property-dao.md)** property of text for string comparisons.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="1b071-142">Constante</span><span class="sxs-lookup"><span data-stu-id="1b071-142">Constant</span></span></p></th>
<th><p><span data-ttu-id="1b071-143">Ordre de classement</span><span class="sxs-lookup"><span data-stu-id="1b071-143">Collating order</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="1b071-144"><strong>dbLangGeneral</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-144"><strong>dbLangGeneral</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-145">Anglais, allemand, français, portugais, italien et espagnol</span><span class="sxs-lookup"><span data-stu-id="1b071-145">English, German, French, Portuguese, Italian, and Modern Spanish</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-146"><strong>dbLangArabic</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-146"><strong>dbLangArabic</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-147">Arabe</span><span class="sxs-lookup"><span data-stu-id="1b071-147">Arabic</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-148"><strong>dbLangChineseSimplified</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-148"><strong>dbLangChineseSimplified</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-149">Chinois simplifié</span><span class="sxs-lookup"><span data-stu-id="1b071-149">Simplified Chinese</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-150"><strong>dbLangChineseTraditional</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-150"><strong>dbLangChineseTraditional</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-151">Chinois traditionnel</span><span class="sxs-lookup"><span data-stu-id="1b071-151">Traditional Chinese</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-152"><strong>dbLangCyrillic</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-152"><strong>dbLangCyrillic</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-153">Russe</span><span class="sxs-lookup"><span data-stu-id="1b071-153">Russian</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-154"><strong>dbLangCzech</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-154"><strong>dbLangCzech</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-155">Tchèque</span><span class="sxs-lookup"><span data-stu-id="1b071-155">Czech</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-156"><strong>dbLangDutch</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-156"><strong>dbLangDutch</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-157">Néerlandais</span><span class="sxs-lookup"><span data-stu-id="1b071-157">Dutch</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-158"><strong>dbLangGreek</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-158"><strong>dbLangGreek</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-159">Grec</span><span class="sxs-lookup"><span data-stu-id="1b071-159">Greek</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-160"><strong>dbLangHebrew</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-160"><strong>dbLangHebrew</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-161">Hébreu</span><span class="sxs-lookup"><span data-stu-id="1b071-161">Hebrew</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-162"><strong>dbLangHungarian</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-162"><strong>dbLangHungarian</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-163">Hongrois</span><span class="sxs-lookup"><span data-stu-id="1b071-163">Hungarian</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-164"><strong>dbLangIcelandic</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-164"><strong>dbLangIcelandic</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-165">Islandais</span><span class="sxs-lookup"><span data-stu-id="1b071-165">Icelandic</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-166"><strong>dbLangJapanese</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-166"><strong>dbLangJapanese</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-167">Japonais</span><span class="sxs-lookup"><span data-stu-id="1b071-167">Japanese</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-168"><strong>dbLangKorean</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-168"><strong>dbLangKorean</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-169">Coréen</span><span class="sxs-lookup"><span data-stu-id="1b071-169">Korean</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-170"><strong>dbLangNordic</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-170"><strong>dbLangNordic</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-171">Langues nordiques (Moteur de base de données Microsoft Jet version 1.0 uniquement)</span><span class="sxs-lookup"><span data-stu-id="1b071-171">Nordic languages (Microsoft Jet database engine version 1.0 only)</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-172"><strong>dbLangNorwDan</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-172"><strong>dbLangNorwDan</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-173">Norvégien et danois</span><span class="sxs-lookup"><span data-stu-id="1b071-173">Norwegian and Danish</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-174"><strong>dbLangPolish</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-174"><strong>dbLangPolish</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-175">Polonais</span><span class="sxs-lookup"><span data-stu-id="1b071-175">Polish</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-176"><strong>dbLangSlovenian</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-176"><strong>dbLangSlovenian</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-177">Slovène</span><span class="sxs-lookup"><span data-stu-id="1b071-177">Slovenian</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-178"><strong>dbLangSpanish</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-178"><strong>dbLangSpanish</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-179">Espagnol traditionnel</span><span class="sxs-lookup"><span data-stu-id="1b071-179">Traditional Spanish</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-180"><strong>dbLangSwedFin</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-180"><strong>dbLangSwedFin</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-181">Suédois et finnois</span><span class="sxs-lookup"><span data-stu-id="1b071-181">Swedish and Finnish</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-182"><strong>dbLangThai</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-182"><strong>dbLangThai</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-183">Thaï</span><span class="sxs-lookup"><span data-stu-id="1b071-183">Thai</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-184"><strong>dbLangTurkish</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-184"><strong>dbLangTurkish</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-185">Turc</span><span class="sxs-lookup"><span data-stu-id="1b071-185">Turkish</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="1b071-186">Vous pouvez utiliser une ou plusieurs des constantes suivantes dans l'argument options pour indiquer la version du format de données et préciser s'il faut chiffrer ou non la base de données.</span><span class="sxs-lookup"><span data-stu-id="1b071-186">You can use one or more of the following constants in the options argument to specify which version the data format should have and whether or not to encrypt the database.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="1b071-187">Constant</span><span class="sxs-lookup"><span data-stu-id="1b071-187">Constant</span></span></p></th>
<th><p><span data-ttu-id="1b071-188">Description</span><span class="sxs-lookup"><span data-stu-id="1b071-188">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="1b071-189"><strong>dbEncrypt</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-189"><strong>dbEncrypt</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-190">Crée une base de données chiffrée.</span><span class="sxs-lookup"><span data-stu-id="1b071-190">Creates an encrypted database.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-191"><strong>dbVersion10</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-191"><strong>dbVersion10</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-192">Crée une base de données qui utilise le format de fichier de la version 1.0 du moteur de base de données Microsoft Jet.</span><span class="sxs-lookup"><span data-stu-id="1b071-192">Creates a database that uses the Microsoft Jet database engine version 1.0 file format.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-193"><strong>dbVersion11</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-193"><strong>dbVersion11</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-194">Crée une base de données qui utilise le format de fichier de la version 1.1 du moteur de base de données Microsoft Jet.</span><span class="sxs-lookup"><span data-stu-id="1b071-194">Creates a database that uses the Microsoft Jet database engine version 1.1 file format.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-195"><strong>dbVersion20</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-195"><strong>dbVersion20</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-196">Crée une base de données qui utilise le format de fichier de la version 2.0 du moteur de base de données Microsoft Jet.</span><span class="sxs-lookup"><span data-stu-id="1b071-196">Creates a database that uses the Microsoft Jet database engine version 2.0 file format.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-197"><strong>dbVersion30</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-197"><strong>dbVersion30</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-198">Crée une base de données qui utilise le format de fichier de la version 3.0 (compatible avec la version 3.5) du moteur de base de données Microsoft Jet.</span><span class="sxs-lookup"><span data-stu-id="1b071-198">Creates a database that uses the Microsoft Jet database engine version 3.0 file format (compatible with version 3.5).</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="1b071-199"><strong>dbVersion40</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-199"><strong>dbVersion40</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-200">Crée une base de données qui utilise le format de fichier de la version 4.0 du moteur de base de données Microsoft Jet.</span><span class="sxs-lookup"><span data-stu-id="1b071-200">Creates a database that uses the Microsoft Jet database engine version 4.0 file format.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="1b071-201"><strong>dbVersion120</strong></span><span class="sxs-lookup"><span data-stu-id="1b071-201"><strong>dbVersion120</strong></span></span></p></td>
<td><p><span data-ttu-id="1b071-202">Crée une base de données qui utilise le format de fichier de la version 12.0 du moteur de base de données Microsoft Access.</span><span class="sxs-lookup"><span data-stu-id="1b071-202">Creates a database that uses the Microsoft Access database engine version 12.0 file format.</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="1b071-203">Si vous omettez la constante de chiffrement, **CreateDatabase** crée une base de données non chiffrée.</span><span class="sxs-lookup"><span data-stu-id="1b071-203">If you omit the encryption constant, **CreateDatabase** creates an un-encrypted database.</span></span>

<span data-ttu-id="1b071-p106">Utilisez la méthode **CreateDatabase** pour créer et ouvrir une nouvelle base de données vide et renvoyer l'objet **Database**. Vous devez compléter sa structure et son contenu en utilisant des objets DAO supplémentaires. Si vous souhaitez réaliser une copie partielle ou complète d'une base de données existante, vous pouvez faire appel à la méthode **[CompactDatabase](dbengine-compactdatabase-method-dao.md)** afin de créer une copie personnalisable.</span><span class="sxs-lookup"><span data-stu-id="1b071-p106">Use the **CreateDatabase** method to create and open a new, empty database, and return the **Database** object. You must complete its structure and content by using additional DAO objects. If you want to make a partial or complete copy of an existing database, you can use the **[CompactDatabase](dbengine-compactdatabase-method-dao.md)** method to make a copy that you can customize.</span></span>
