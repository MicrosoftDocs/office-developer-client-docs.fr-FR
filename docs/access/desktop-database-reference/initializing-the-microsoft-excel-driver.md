---
title: Initialisation du pilote Microsoft Excel
TOCTitle: Initializing the Microsoft Excel Driver
ms:assetid: 06c7f823-8e74-0811-cc00-e6b32075ef11
ms:mtpsurl: https://msdn.microsoft.com/library/Ff844939(v=office.15)
ms:contentKeyID: 48543054
ms.date: 09/18/2015
mtps_version: v=office.15
f1_keywords:
- acmain11.chm1032159
f1_categories:
- Office.Version=v15
ms.openlocfilehash: c79d859b122eb3595c31b2ffcec192e2d69ed7b4
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472158"
---
# <a name="initializing-the-microsoft-excel-driver"></a><span data-ttu-id="11a86-102">Initialisation du pilote Microsoft Excel</span><span class="sxs-lookup"><span data-stu-id="11a86-102">Initializing the Microsoft Excel Driver</span></span>


<span data-ttu-id="11a86-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="11a86-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="11a86-p101">Lorsque vous installez le pilote de base de données Microsoft® Excel, le programme d'installation écrit un ensemble de valeurs par défaut dans les sous-clé Engines et ISAM Formats du registre Microsoft Windows®. Vous ne devez pas modifier directement ces paramètres. Utilisez le programme d'installation de votre application pour ajouter, supprimer ou modifier ces paramètres. Les sections suivantes décrivent l'initialisation et les paramètres de format ISAM pour le pilote de base de données Microsoft Excel.</span><span class="sxs-lookup"><span data-stu-id="11a86-p101">When you install the Microsoft® Excel driver, the Setup program writes a set of default values to the Microsoft Windows® Registry in the Engines and ISAM Formats subkeys. You should not modify these settings directly; use the setup program for your application to add, remove, or change these settings. The following sections describe initialization and ISAM Format settings for the Microsoft Excel database driver.</span></span>

## <a name="microsoft-excel-initialization-settings"></a><span data-ttu-id="11a86-107">Paramètres d'initialisation Microsoft Excel</span><span class="sxs-lookup"><span data-stu-id="11a86-107">Microsoft Excel Initialization Settings</span></span>

<span data-ttu-id="11a86-108">Le **Access Connectivity Engine\\moteurs\\Excel** dossier contient des paramètres d’initialisation du pilote Aceexcl.dll, utilisé pour l’accès externe aux feuilles de calcul Microsoft Excel.</span><span class="sxs-lookup"><span data-stu-id="11a86-108">The **Access Connectivity Engine\\Engines\\Excel** folder includes initialization settings for the Aceexcl.dll driver, used for external access to Microsoft Excel worksheets.</span></span> <span data-ttu-id="11a86-109">L'exemple ci-après montre des paramètres par défaut pour les entrées de ce dossier.</span><span class="sxs-lookup"><span data-stu-id="11a86-109">Typical settings for the entries in this folder are shown in the following example.</span></span>

```vb
    win32=<path>\ Aceexcl.dll  
    
    TypeGuessRows=8 
    
    ImportMixedTypes=Text 
    
    AppendBlankRows=1 
    
    FirstRowHasNames=Yes
```

<span data-ttu-id="11a86-110">Le moteur de base de données Microsoft Access utilise les entrées de dossier Excel suivantes.</span><span class="sxs-lookup"><span data-stu-id="11a86-110">The Microsoft Access database engine uses the Excel folder entries as follows.</span></span>

<table>
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="11a86-111">Entrée</span><span class="sxs-lookup"><span data-stu-id="11a86-111">Entry</span></span></p></th>
<th><p><span data-ttu-id="11a86-112">Description</span><span class="sxs-lookup"><span data-stu-id="11a86-112">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="11a86-113">Win32</span><span class="sxs-lookup"><span data-stu-id="11a86-113">win32</span></span></p></td>
<td><p><span data-ttu-id="11a86-p103">Emplacement de msexcl40.dll. Le Chemin d'accès complet est déterminé au moment de l'installation. Les valeurs sont de type REG_SZ.</span><span class="sxs-lookup"><span data-stu-id="11a86-p103">The location of msexcl40.dll. The full path is determined at the time of installation. Values are of type REG_SZ.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="11a86-117">TypeGuessRows</span><span class="sxs-lookup"><span data-stu-id="11a86-117">TypeGuessRows</span></span></p></td>
<td><p><span data-ttu-id="11a86-118">Le nombre de lignes à vérifier pour le type de données.</span><span class="sxs-lookup"><span data-stu-id="11a86-118">The number of rows to be checked for the data type.</span></span> <span data-ttu-id="11a86-119">Le type de données est déterminé étant donné le nombre maximal de types de données trouvés.</span><span class="sxs-lookup"><span data-stu-id="11a86-119">The data type is determined given the maximum number of kinds of data found.</span></span> <span data-ttu-id="11a86-120">S’il existe un lien, le type de données est déterminé dans l’ordre suivant : nombre, devise, Date, texte, booléen.</span><span class="sxs-lookup"><span data-stu-id="11a86-120">If there is a tie, the data type is determined in the following order: Number, Currency, Date, Text, Boolean.</span></span> <span data-ttu-id="11a86-121">Si les données rencontrées qui ne correspond pas au type de données estimé pour la colonne, elle est renvoyée comme une valeur <strong>Null</strong> .</span><span class="sxs-lookup"><span data-stu-id="11a86-121">If data is encountered that does not match the data type guessed for the column, it is returned as a <strong>Null</strong> value.</span></span> <span data-ttu-id="11a86-122">À l’importation, si une colonne est une combinaison de types de données, la colonne entière sera convertie en fonction du paramètre ImportMixedTypes.</span><span class="sxs-lookup"><span data-stu-id="11a86-122">On import, if a column has mixed data types, the entire column will be cast according to the ImportMixedTypes setting.</span></span> <span data-ttu-id="11a86-123">Le nombre de lignes à vérifier par défaut est 8.</span><span class="sxs-lookup"><span data-stu-id="11a86-123">The default number of rows to be checked is 8.</span></span> <span data-ttu-id="11a86-124">Les valeurs sont de type REG_DWORD.</span><span class="sxs-lookup"><span data-stu-id="11a86-124">Values are of type REG_DWORD.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="11a86-125">ImportMixedTypes</span><span class="sxs-lookup"><span data-stu-id="11a86-125">ImportMixedTypes</span></span></p></td>
<td><p><span data-ttu-id="11a86-p105">Peut être défini avec la valeur MajorityType ou Text. Avec MajorityType, les colonnes de types de données mélangés sont mises en forme en fonction du type de données prédominant à l'importation. Avec Text, elles seront mises en forme selon le type de données texte à l'importation. La valeur par défaut est Text. Les values sont de type REG_SZ.</span><span class="sxs-lookup"><span data-stu-id="11a86-p105">Can be set to MajorityType or Text. If set to MajorityType, columns of mixed data types will be cast to the predominate data type on import. If set to Text, columns of mixed data types will be cast to Text on import. The default is Text. Values are of type REG_SZ.</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="11a86-131">AppendBlankRows possède</span><span class="sxs-lookup"><span data-stu-id="11a86-131">AppendBlankRows</span></span></p></td>
<td><p><span data-ttu-id="11a86-p106">Nombre de lignes vides devant être ajouté à la fin d'une feuille de calcul version 3.5 ou 4.0, avant l'ajout de nouvelles données. Par exemple, si AppendBlankRows possède la valeur 4, Microsoft Jet ajoutera 4 lignes vides à la fin de la feuille de calcul avant d'ajouter les lignes qui contiennent des données. Les valeurs entières pour ce paramètre peuvent être comprises entre 0 entre 16 ; la valeur par défaut est 01 (une ligne supplémentaire). Les valeurs sont de type REG_DWORD.</span><span class="sxs-lookup"><span data-stu-id="11a86-p106">The number of blank rows to be appended to the end of a Version 3.5 or Version 4.0 worksheet before new data is added. For example, if AppendBlankRows is set to 4, Microsoft Jet will append 4 blank rows to the end of the worksheet before appending rows that contain data. Integer values for this setting can range from 0 to 16; the default is 01 (one additional row appended). Values are of type REG_DWORD.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="11a86-136">FirstRowHasNames</span><span class="sxs-lookup"><span data-stu-id="11a86-136">FirstRowHasNames</span></span></p></td>
<td><p><span data-ttu-id="11a86-p107">Valeur binaire qui indique si la première ligne du tableau contient des noms de colonne. La valeur 01 indique que les noms de colonne sont pris de la première ligne, pendant l'importation. La valeur 00 indique l'absence de nom de colonne dans la première ligne ; les noms de colonne sont alors F1, F2, F3 etc. La valeur par défaut est 01. Les valeurs sont de type REG_BINARY.</span><span class="sxs-lookup"><span data-stu-id="11a86-p107">A binary value that indicates whether the first row of the table contains column names. A value of 01 indicates that, during import, column names are taken from the first row. A value of 00 indicates no column names in the first row; column names appear as F1, F2, F3, and so on. The default is 01. Values are of type REG_BINARY.</span></span></p></td>
</tr>
</tbody>
</table>


<span data-ttu-id="11a86-142">Le **Access Connectivity Engine\\moteurs\\Excel 8.0** dossier contient les entrées suivantes, qui s’appliquent à Microsoft Excel 97.</span><span class="sxs-lookup"><span data-stu-id="11a86-142">The **Access Connectivity Engine\\Engines\\Excel 8.0** folder contains the following entries, which apply to Microsoft Excel 97.</span></span>

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="11a86-143">Nom d'entrée</span><span class="sxs-lookup"><span data-stu-id="11a86-143">Entry name</span></span></p></th>
<th><p><span data-ttu-id="11a86-144">Type</span><span class="sxs-lookup"><span data-stu-id="11a86-144">Type</span></span></p></th>
<th><p><span data-ttu-id="11a86-145">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a86-145">Value</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="11a86-146">Engine</span><span class="sxs-lookup"><span data-stu-id="11a86-146">Engine</span></span></p></td>
<td><p><span data-ttu-id="11a86-147">REG_SZ</span><span class="sxs-lookup"><span data-stu-id="11a86-147">REG_SZ</span></span></p></td>
<td><p><span data-ttu-id="11a86-148">Excel</span><span class="sxs-lookup"><span data-stu-id="11a86-148">Excel</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="11a86-149">ExportFilter</span><span class="sxs-lookup"><span data-stu-id="11a86-149">ExportFilter</span></span></p></td>
<td><p><span data-ttu-id="11a86-150">REG_SZ</span><span class="sxs-lookup"><span data-stu-id="11a86-150">REG_SZ</span></span></p></td>
<td><p><span data-ttu-id="11a86-151">Microsoft Excel 97-2000 (\*.xls)</span><span class="sxs-lookup"><span data-stu-id="11a86-151">Microsoft Excel 97-2000 (\*.xls)</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="11a86-152">CanLink</span><span class="sxs-lookup"><span data-stu-id="11a86-152">CanLink</span></span></p></td>
<td><p><span data-ttu-id="11a86-153">REG_BINARY</span><span class="sxs-lookup"><span data-stu-id="11a86-153">REG_BINARY</span></span></p></td>
<td><p><span data-ttu-id="11a86-154">01</span><span class="sxs-lookup"><span data-stu-id="11a86-154">01</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="11a86-155">OneTablePerFile</span><span class="sxs-lookup"><span data-stu-id="11a86-155">OneTablePerFile</span></span></p></td>
<td><p><span data-ttu-id="11a86-156">REG_BINARY</span><span class="sxs-lookup"><span data-stu-id="11a86-156">REG_BINARY</span></span></p></td>
<td><p><span data-ttu-id="11a86-157">00</span><span class="sxs-lookup"><span data-stu-id="11a86-157">00</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="11a86-158">IsamType</span><span class="sxs-lookup"><span data-stu-id="11a86-158">IsamType</span></span></p></td>
<td><p><span data-ttu-id="11a86-159">REG_DWORD</span><span class="sxs-lookup"><span data-stu-id="11a86-159">REG_DWORD</span></span></p></td>
<td><p><span data-ttu-id="11a86-160">1</span><span class="sxs-lookup"><span data-stu-id="11a86-160">1</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="11a86-161">IndexDialog</span><span class="sxs-lookup"><span data-stu-id="11a86-161">IndexDialog</span></span></p></td>
<td><p><span data-ttu-id="11a86-162">REG_BINARY</span><span class="sxs-lookup"><span data-stu-id="11a86-162">REG_BINARY</span></span></p></td>
<td><p><span data-ttu-id="11a86-163">00</span><span class="sxs-lookup"><span data-stu-id="11a86-163">00</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="11a86-164">CreateDBOnExport</span><span class="sxs-lookup"><span data-stu-id="11a86-164">CreateDBOnExport</span></span></p></td>
<td><p><span data-ttu-id="11a86-165">REG_BINARY</span><span class="sxs-lookup"><span data-stu-id="11a86-165">REG_BINARY</span></span></p></td>
<td><p><span data-ttu-id="11a86-166">01</span><span class="sxs-lookup"><span data-stu-id="11a86-166">01</span></span></p></td>
</tr>
<tr class="even">
<td><p><span data-ttu-id="11a86-167">ResultTextExport</span><span class="sxs-lookup"><span data-stu-id="11a86-167">ResultTextExport</span></span></p></td>
<td><p><span data-ttu-id="11a86-168">REG_SZ</span><span class="sxs-lookup"><span data-stu-id="11a86-168">REG_SZ</span></span></p></td>
<td><p><span data-ttu-id="11a86-p108">Permet d'exporter les données de la base de données active vers le fichier Microsoft Excel 97. Ce processus entraîne l'écrasement des données si l'exportation a lieu vers un fichier existant.</span><span class="sxs-lookup"><span data-stu-id="11a86-p108">Export data from the current database into a Microsoft Excel 97 file. This process will overwrite the data if exported to an existing file.</span></span></p></td>
</tr>
<tr class="odd">
<td><p><span data-ttu-id="11a86-171">SupportsLongNames</span><span class="sxs-lookup"><span data-stu-id="11a86-171">SupportsLongNames</span></span></p></td>
<td><p><span data-ttu-id="11a86-172">REG_BINARY</span><span class="sxs-lookup"><span data-stu-id="11a86-172">REG_BINARY</span></span></p></td>
<td><p><span data-ttu-id="11a86-173">01</span><span class="sxs-lookup"><span data-stu-id="11a86-173">01</span></span></p></td>
</tr>
</tbody>
</table>



> [!NOTE]
> <span data-ttu-id="11a86-174">Lorsque vous modifiez des paramètres de registre Windows, vous devez redémarrer le moteur de base de données pour que les nouveaux paramètres entrent en vigueur.</span><span class="sxs-lookup"><span data-stu-id="11a86-174">When you change Windows Registry settings, you must exit and then restart the database engine for the new settings to take effect.</span></span>

