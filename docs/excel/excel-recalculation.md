---
title: Recalcul Excel
manager: kelbow
ms.date: 08/22/2018
ms.audience: Developer
ms.topic: overview
keywords:
- forced calculation [excel 2007],selective recalculation [Excel 2007],functions [Excel 2007], volatile,calculation modes,recalculated cells [Excel 2007],dependence [Excel 2007],specified worksheet calculation [Excel 2007],recalculation [Excel 2007],workbook tree rebuild [Excel 2007],range calculation [Excel 2007],Excel recalculation,volatile functions [Excel 2007],functions [Excel 2007], non-volatile,active worksheet calculation [Excel 2007],dirty cells [Excel 2007],non-volatile functions [Excel 2007],forced recalculation [Excel 2007]
ms.assetid: b4c38442-42e6-4fd2-a1b0-97cfa3300379
description: 'S’applique à : Excel 2013 | Office 2013 | Visual Studio'
localization_priority: Priority
ms.openlocfilehash: c4b6dbad7a31b7155d1bec3a5c867b6c74d0ff42
ms.sourcegitcommit: 35b723efe168ae4bad461bd16b26f9a2412656f2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2021
ms.locfileid: "53139092"
---
# <a name="excel-recalculation"></a><span data-ttu-id="908ff-104">Recalcul Excel</span><span class="sxs-lookup"><span data-stu-id="908ff-104">Excel Recalculation</span></span>

<span data-ttu-id="908ff-105">**S’applique à**: Excel 2013 | Office 2013 | Visual Studio</span><span class="sxs-lookup"><span data-stu-id="908ff-105">**Applies to**: Excel 2013 | Office 2013 | Visual Studio</span></span> 
  
<span data-ttu-id="908ff-106">L’utilisateur peut déclencher le recalcul dans Microsoft Excel de plusieurs façons :</span><span class="sxs-lookup"><span data-stu-id="908ff-106">The user can trigger recalculation in Microsoft Excel in several ways, for example:</span></span>
  
- <span data-ttu-id="908ff-107">en saisissant de nouvelles données (si Excel est en mode de recalcul automatique, décrit plus loin dans cette rubrique) ;</span><span class="sxs-lookup"><span data-stu-id="908ff-107">Entering new data (if Excel is in Automatic recalculation mode, described later in this topic).</span></span>
    
- <span data-ttu-id="908ff-108">en demandant explicitement à Excel de recalculer l’intégralité ou la partie d’un classeur ;</span><span class="sxs-lookup"><span data-stu-id="908ff-108">Explicitly instructing Excel to recalculate all or part of a workbook.</span></span>
    
- <span data-ttu-id="908ff-109">en supprimant ou insérant une ligne ou colonne ;</span><span class="sxs-lookup"><span data-stu-id="908ff-109">Deleting or inserting a row or column.</span></span>
    
- <span data-ttu-id="908ff-110">en enregistrant un classeur alors que l’option **Recalcul avant l’enregistrement** est définie ;</span><span class="sxs-lookup"><span data-stu-id="908ff-110">Saving a workbook while the **Recalculate before save** option is set.</span></span> 
    
- <span data-ttu-id="908ff-111">en exécutant certaines actions de filtrage automatique ;</span><span class="sxs-lookup"><span data-stu-id="908ff-111">Performing certain Autofilter actions.</span></span>
    
- <span data-ttu-id="908ff-112">en cliquant deux fois sur un séparateur de ligne ou de colonne (en mode de calcul automatique) ;</span><span class="sxs-lookup"><span data-stu-id="908ff-112">Double-clicking a row or column divider (in Automatic calculation mode).</span></span>
    
- <span data-ttu-id="908ff-113">en ajoutant, modifiant ou supprimant un nom défini ;</span><span class="sxs-lookup"><span data-stu-id="908ff-113">Adding, editing, or deleting a defined name.</span></span>
    
- <span data-ttu-id="908ff-114">en renommant une feuille de calcul ;</span><span class="sxs-lookup"><span data-stu-id="908ff-114">Renaming a worksheet.</span></span>
    
- <span data-ttu-id="908ff-115">en modifiant la position d’une feuille de calcul liée à d’autres feuilles de calcul ;</span><span class="sxs-lookup"><span data-stu-id="908ff-115">Changing the position of a worksheet in relation to other worksheets.</span></span>
    
- <span data-ttu-id="908ff-116">en masquant ou affichant des lignes, mais pas des colonnes.</span><span class="sxs-lookup"><span data-stu-id="908ff-116">Hiding or unhiding rows, but not columns.</span></span>
    
> [!NOTE]
> <span data-ttu-id="908ff-p101">Cette rubrique ne fait pas la différence entre l’utilisateur appuyant sur une touche ou cliquant sur la souris, et les tâches effectuées par une commande ou une macro. L’utilisateur exécute la commande, ou fait quelque chose pour déclencher l’exécution de la commande, de façon à ce que ce soit considéré comme une action de l’utilisateur. C’est pourquoi les termes « l’utilisateur » signifient également « l’utilisateur, ou une commande ou processus exécuté par l’utilisateur ».</span><span class="sxs-lookup"><span data-stu-id="908ff-p101">This topic does not distinguish between the user directly pressing a key or clicking the mouse, and those tasks being done by a command or macro. The user runs the command, or does something to cause the command to run so that it is still considered a user action. Therefore the phrase "the user" also means "the user, or a command or process started by the user."</span></span> 
  
## <a name="dependence-dirty-cells-and-recalculated-cells"></a><span data-ttu-id="908ff-120">Dépendance, cellules incorrectes et cellules recalculées</span><span class="sxs-lookup"><span data-stu-id="908ff-120">Dependence, Dirty Cells, and Recalculated Cells</span></span>

<span data-ttu-id="908ff-121">Le calcul des feuilles de calcul dans Excel peut être considéré comme un processus à trois étapes :</span><span class="sxs-lookup"><span data-stu-id="908ff-121">The calculation of worksheets in Excel can be viewed as a three-stage process:</span></span>
  
1. <span data-ttu-id="908ff-122">Construction d’une arborescence des dépendances</span><span class="sxs-lookup"><span data-stu-id="908ff-122">Construction of a dependency tree</span></span>
    
2. <span data-ttu-id="908ff-123">Construction d’une chaîne de calcul</span><span class="sxs-lookup"><span data-stu-id="908ff-123">Construction of a calculation chain</span></span>
    
3. <span data-ttu-id="908ff-124">Recalcul de cellules</span><span class="sxs-lookup"><span data-stu-id="908ff-124">Recalculation of cells</span></span>
    
<span data-ttu-id="908ff-p102">L’arborescence des dépendances indique à Excel les cellules dépendant d’autres cellules, ou de façon équivalente, les cellules qui sont les antécédents d’autres cellules. À partir de cette arborescence, Excel crée une chaîne de calcul. Celle-ci répertorie toutes les cellules qui contiennent des formules dans l’ordre dans lequel elles doivent être calculées. Pendant le recalcul, Excel modifie cette chaîne s’il rencontre une formule qui dépend d’une cellule qui n’a pas encore été calculée. Dans ce cas, la cellule est calculée, et les cellules qui en dépendent sont déplacées vers le bas de la chaîne. Pour cette raison, les durées de calcul peuvent souvent être réduites dans une feuille de calcul ayant juste été ouverte lors des premiers cycles de calcul.</span><span class="sxs-lookup"><span data-stu-id="908ff-p102">The dependency tree informs Excel about which cells depend on which others, or equivalently, which cells are precedents for which others. From this tree, Excel constructs a calculation chain. The calculation chain lists all the cells that contain formulas in the order in which they should be calculated. During recalculation, Excel revises this chain if it comes across a formula that depends on a cell that has not yet been calculated. In this case, the cell that is being calculated and its dependents are moved down the chain. For this reason, calculation times can often improve in a worksheet that has just been opened in the first few calculation cycles.</span></span>
  
<span data-ttu-id="908ff-p103">Lorsqu’une modification structurelle est apportée à un classeur, par exemple lorsqu’une nouvelle formule est saisie, Excel reconstruit l’arborescence des dépendances et la chaîne de calcul. Lors de la saisie de nouvelles données ou formules, Excel marque toutes les cellules qui dépendent de ces nouvelles données comme nécessitant un recalcul. Les cellules marquées de cette manière sont appelées *incorrectes*. Toutes les cellules directement et indirectement dépendantes sont marquées comme incorrectes. De ce fait, si B1 dépend d’A1 et que C1 dépend de B1, lorsque A1 est modifiée, B1 et C1 sont marquées comme incorrectes.</span><span class="sxs-lookup"><span data-stu-id="908ff-p103">When a structural change is made to a workbook, for example, when a new formula is entered, Excel reconstructs the dependency tree and calculation chain. When new data or new formulas are entered, Excel marks all the cells that depend on that new data as needing recalculation. Cells that are marked in this way are known as  *dirty*  . All direct and indirect dependents are marked as dirty so that if B1 depends on A1, and C1 depends on B1, when A1 is changed, both B1 and C1 are marked as dirty.</span></span> 
  
<span data-ttu-id="908ff-p104">Si une cellule dépend directement ou indirectement d’elle-même, Excel détecte la référence circulaire et avertit l’utilisateur. Il s’agit généralement d’une condition d’erreur que l’utilisateur doit résoudre, et Excel fournit des outils de graphique et de navigation très utiles pour aider l’utilisateur à trouver la source de la dépendance circulaire. Dans certains cas, vous voudrez délibérément que cette condition existe. Par exemple, vous souhaiterez peut-être exécuter un calcul itératif où le point de départ pour l’itération suivante est le résultat de l’itération précédente. Excel prend en charge le contrôle des calculs itératifs via la boîte de dialogue des options de calcul.</span><span class="sxs-lookup"><span data-stu-id="908ff-p104">If a cell depends, directly or indirectly, on itself, Excel detects the circular reference and warns the user. This is usually an error condition that the user must fix, and Excel provides very helpful graphical and navigational tools to help the user to find the source of the circular dependency. In some cases, you might deliberately want this condition to exist. For example, you might want to run an iterative calculation where the starting point for the next iteration is the result of the previous iteration. Excel supports control of iterative calculations through the calculation options dialog box.</span></span>
  
<span data-ttu-id="908ff-p105">Une fois que certaines cellules sont marquées comme incorrectes, lorsqu’un recalcul est effectué, Excel réévalue le contenu de chaque cellule incorrecte dans l’ordre déterminé par la chaîne de calcul. Dans l’exemple donné précédemment, cela signifie que B1 est calculée en premier lieu, suivie par C1. Ce calcul a lieu immédiatement après qu’Excel a terminé de marquer les cellules comme incorrectes si le mode de recalcul est automatique. Sinon, il se produit plus tard.</span><span class="sxs-lookup"><span data-stu-id="908ff-p105">After marking cells as dirty, when a recalculation is next done, Excel reevaluates the contents of each dirty cell in the order dictated by the calculation chain. In the example given earlier, this means B1 is first, and then C1. This recalculation occurs immediately after Excel finishes marking cells as dirty if the recalculation mode is automatic; otherwise, it occurs later.</span></span>
  
<span data-ttu-id="908ff-p106">À partir de Microsoft Excel 2002, l’objet **Range** dans Microsoft Visual Basic pour Applications (VBA) prend en charge une méthode, **Range.Dirty**, qui indique les cellules nécessitant un calcul. Lorsqu’elle est utilisée conjointement à la méthode **Range.Calculate** (voir la section suivante), elle permet le recalcul forcé de cellules dans une plage donnée. Cela est utile lorsque vous effectuez un calcul limité au cours d’une macro, où le mode de calcul est défini sur manuel, afin d’éviter la surcharge de calcul de cellules non liées à la fonction de macro. Les méthodes de calcul de plage ne sont pas disponibles via l’API C.</span><span class="sxs-lookup"><span data-stu-id="908ff-p106">Starting in Microsoft Excel 2002, the **Range** object in Microsoft Visual Basic for Applications (VBA) supports a method, **Range.Dirty**, which marks cells as needing calculation. When it is used together with the **Range.Calculate** method (see next section), it enables forced recalculation of cells in a given range. This is useful when you are performing a limited calculation during a macro, where the calculation mode is set to manual, to avoid the overhead of calculating cells unrelated to the macro function. Range calculation methods are not available through the C API.</span></span> 
  
<span data-ttu-id="908ff-p107">Dans Excel 2002 et versions antérieures, Excel a construit une chaîne de calcul pour chaque feuille de calcul dans chaque classeur ouvert. Cela a généré une certaine complexité dans la gestion des liens entre les feuilles de calcul, et a demandé de l’attention pour garantir un recalcul efficace. En particulier, dans Excel 2000, vous devez réduire les dépendances entre les classeurs et nommer les feuilles de calcul par ordre alphabétique, afin que les feuilles dépendant d’autres feuilles apparaissent après les feuilles dont elles dépendent.</span><span class="sxs-lookup"><span data-stu-id="908ff-p107">In Excel 2002 and earlier versions, Excel built a calculation chain for each worksheet in each open workbook. This resulted in some complexity in the way links between worksheets were handled, and required some care to ensure efficient recalculation. In particular, in Excel 2000, you should minimize cross-worksheet dependencies and name worksheets in alphabetical order so that sheets that depend on other sheets come alphabetically after the sheets they depend on.</span></span>
  
<span data-ttu-id="908ff-p108">Dans Excel 2007, la logique a été améliorée pour permettre le recalcul sur plusieurs threads afin que les sections de la chaîne de calcul ne soient pas interdépendantes et puissent être calculées en même temps. Vous pouvez configurer Excel pour utiliser plusieurs threads sur un ordinateur monoprocesseur ou un thread unique sur un ordinateur multiprocesseur ou multicœur.</span><span class="sxs-lookup"><span data-stu-id="908ff-p108">In Excel 2007, the logic was improved to enable recalculation on multiple threads so that sections of the calculation chain are not interdependent and can be calculated at the same time. You can configure Excel to use multiple threads on a single processor computer, or a single thread on a multi-processor or multi-core computer.</span></span> 
  
## <a name="asynchronous-user-defined-functions-udfs"></a><span data-ttu-id="908ff-152">Fonctions définies par l’utilisateur (UDF) asynchrones</span><span class="sxs-lookup"><span data-stu-id="908ff-152">Asynchronous User Defined Functions (UDFs)</span></span>

<span data-ttu-id="908ff-p109">Lorsqu’un calcul rencontre une UDF asynchrone, il enregistre l’état de la formule en cours, démarre l’UDF et continue l’évaluation des autres cellules. Lorsque le calcul termine l’évaluation des cellules, Excel attend que les fonctions asynchrones se terminent s’il en existe encore en cours d’exécution. Lorsque les fonctions asynchrones génèrent des résultats, Excel termine la formule et exécute une nouvelle phase de calcul pour recalculer les cellules qui utilisent la cellule référencée à la fonction asynchrone.</span><span class="sxs-lookup"><span data-stu-id="908ff-p109">When a calculation encounters an asynchronous UDF, it saves the state of the current formula, starts the UDF and continues evaluating the rest of the cells. When the calculation finishes evaluating the cells Excel waits for the asynchronous functions to complete if there are still asynchronous functions running. As each asynchronous function reports results, Excel finishes the formula, and then runs a new calculation pass to re-compute cells that use the cell with the reference to the asynchronous function.</span></span>
  
## <a name="volatile-and-non-volatile-functions"></a><span data-ttu-id="908ff-156">Fonctions volatiles et non volatiles</span><span class="sxs-lookup"><span data-stu-id="908ff-156">Volatile and Non-Volatile Functions</span></span>

<span data-ttu-id="908ff-p110">Excel prend en charge le concept d’une fonction volatile, c’est-à-dire d’une fonction dont la valeur ne peut pas rester la même d’un moment à l’autre, même si aucun de ses arguments (si elle en a) n’a été modifié. Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes les cellules dépendantes, chaque fois qu’il réalise un recalcul. C’est pourquoi un nombre trop élevé de fonctions volatiles peut ralentir le recalcul. Utilisez-les avec parcimonie.</span><span class="sxs-lookup"><span data-stu-id="908ff-p110">Excel supports the concept of a volatile function, that is, one whose value cannot be assumed to be the same from one moment to the next even if none of its arguments (if it takes any) has changed. Excel reevaluates cells that contain volatile functions, together with all dependents, every time that it recalculates. For this reason, too much reliance on volatile functions can make recalculation times slow. Use them sparingly.</span></span>
  
<span data-ttu-id="908ff-161">Les fonctions Excel suivantes sont volatiles :</span><span class="sxs-lookup"><span data-stu-id="908ff-161">The following Excel functions are volatile:</span></span>
  
- <span data-ttu-id="908ff-162">**NOW**</span><span class="sxs-lookup"><span data-stu-id="908ff-162">**NOW**</span></span>
    
- <span data-ttu-id="908ff-163">**TODAY**</span><span class="sxs-lookup"><span data-stu-id="908ff-163">**TODAY**</span></span>
    
- <span data-ttu-id="908ff-164">**RANDBETWEEN**</span><span class="sxs-lookup"><span data-stu-id="908ff-164">**RANDBETWEEN**</span></span>
    
- <span data-ttu-id="908ff-165">**OFFSET**</span><span class="sxs-lookup"><span data-stu-id="908ff-165">**OFFSET**</span></span>
    
- <span data-ttu-id="908ff-166">**INDIRECT**</span><span class="sxs-lookup"><span data-stu-id="908ff-166">**INDIRECT**</span></span>
    
- <span data-ttu-id="908ff-167">**INFO** (en fonction de ses arguments)</span><span class="sxs-lookup"><span data-stu-id="908ff-167">**INFO** (depending on its arguments)</span></span> 
    
- <span data-ttu-id="908ff-168">**CELL** (en fonction de ses arguments)</span><span class="sxs-lookup"><span data-stu-id="908ff-168">**CELL** (depending on its arguments)</span></span> 
    
- <span data-ttu-id="908ff-169">**SUMIF** (en fonction de ses arguments)</span><span class="sxs-lookup"><span data-stu-id="908ff-169">**SUMIF** (depending on its arguments)</span></span> 
    
<span data-ttu-id="908ff-p111">Les VBA et l’API C prennent en charge des méthodes pour indiquer à Excel qu’une fonction définie par l’utilisateur (UDF) doit être traitée comme une fonction volatile. À l’aide de VBA, l’UDF est déclarée comme volatile de la manière suivante.</span><span class="sxs-lookup"><span data-stu-id="908ff-p111">Both the VBA and C API support ways to inform Excel that a user-defined function (UDF) should be handled as volatile. By using VBA, the UDF is declared as volatile as follows.</span></span>
  
```vba
Function MyUDF(MakeMeVolatile As Boolean) As Double
   ' Good practice to call this on the first line.
   Application.Volatile MakeMeVolatile
   MyUDF = Now
End Function

```

<span data-ttu-id="908ff-p112">Par défaut, Excel suppose que les UDF VBA ne sont pas volatiles. Excel apprend qu’une UDF est volatile lorsqu’il l’appelle pour la première fois. Une UDF volatile peut être modifiée en non volatile comme dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="908ff-p112">By default, Excel assumes that VBA UDFs are not volatile. Excel only learns that a UDF is volatile when it first calls it. A volatile UDF can be changed back to non-volatile as in this example.</span></span>
  
<span data-ttu-id="908ff-p113">À l’aide de l’API C, vous pouvez enregistrer une fonction XLL comme étant volatile avant son premier appel. Elle permet également d’activer et de désactiver le statut volatile d’une fonction de feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="908ff-p113">Using the C API, you can register an XLL function as volatile before its first call. It also enables you to switch on and off the volatile status of a worksheet function.</span></span>
  
<span data-ttu-id="908ff-p114">Par défaut, Excel gère les UDF XLL qui acceptent des arguments de plage et qui sont déclarées comme feuilles macro, équivalent de volatiles. Vous pouvez désactiver cet état par défaut à l’aide de la fonction **xlfVolatile** lorsque l’UDF est appelée pour la première fois.</span><span class="sxs-lookup"><span data-stu-id="908ff-p114">By default, Excel handles XLL UDFs that take range arguments and that are declared as macro-sheet equivalents as volatile. You can turn this default state off using the **xlfVolatile** function when the UDF is first called.</span></span> 
  
## <a name="calculation-modes-commands-selective-recalculation-and-data-tables"></a><span data-ttu-id="908ff-179">Modes de calcul, commandes, recalcul sélectif et tables de données</span><span class="sxs-lookup"><span data-stu-id="908ff-179">Calculation Modes, Commands, Selective Recalculation, and Data Tables</span></span>

<span data-ttu-id="908ff-180">Excel propose trois modes de calcul :</span><span class="sxs-lookup"><span data-stu-id="908ff-180">Excel has three calculation modes:</span></span>
  
- <span data-ttu-id="908ff-181">Automatic</span><span class="sxs-lookup"><span data-stu-id="908ff-181">Automatic</span></span>
    
- <span data-ttu-id="908ff-182">Automatique sauf dans les tables</span><span class="sxs-lookup"><span data-stu-id="908ff-182">Automatic Except Tables</span></span>
    
- <span data-ttu-id="908ff-183">Manual</span><span class="sxs-lookup"><span data-stu-id="908ff-183">Manual</span></span>
    
<span data-ttu-id="908ff-p115">Lorsque le calcul est défini sur Automatique, le recalcul se produit après chaque entrée de données, et après certains événements tels que les exemples fournis dans la section précédente. Pour les classeurs très volumineux, le temps de recalcul peut être si long que les utilisateurs doivent le limiter lorsque cela se produit, c’est-à-dire les recalculer uniquement lorsque cela est nécessaire. Pour ce faire, Excel prend en charge le mode Manuel. L’utilisateur peut sélectionner le mode via le système de menus Excel ou par programmation à l’aide de VBA, COM ou de l’API C.</span><span class="sxs-lookup"><span data-stu-id="908ff-p115">When calculation is set to automatic, recalculation occurs after every data input and after certain events such as the examples given in the previous section. For very large workbooks, recalculation time might be so long that users must limit when this happens, that is, only recalculating when they need to. To enable this, Excel supports the manual mode. The user can select the mode through the Excel menu system, or programmatically using VBA, COM, or the C API.</span></span>
  
<span data-ttu-id="908ff-p116">Les tables de données sont des structures spéciales dans une feuille de calcul. Tout d’abord, l’utilisateur définit le calcul d’un résultat sur une feuille de calcul. Cela dépend d’une ou de deux entrées modifiables clé et d’autres paramètres. L’utilisateur peut ensuite créer une table de résultats pour un ensemble de valeurs pour l’une des valeurs clés, ou les deux. La table est créée à l’aide de l’**Assistant Table de données**. Une fois la table configurée, Excel transmet les entrées une par une au calcul et copie la valeur résultante dans la table. Étant donné qu’une ou deux entrées peuvent être utilisées, les tables de données peuvent comporter une ou deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="908ff-p116">Data tables are special structures in a worksheet. First, the user sets up the calculation of a result on a worksheet. This depends on one or two key changeable inputs and other parameters. The user can then create a table of results for a set of values for one or both of the key inputs. The table is created by using the **Data Table Wizard**. After the table is set up, Excel plugs the inputs one-by-one into the calculation and copies the resulting value into the table. As one or two inputs can be used, data tables can be one- or two-dimensional.</span></span> 
  
<span data-ttu-id="908ff-195">Le recalcul des tables de données est légèrement différent :</span><span class="sxs-lookup"><span data-stu-id="908ff-195">Recalculation of data tables is handled slightly differently:</span></span>
  
- <span data-ttu-id="908ff-196">Le recalcul est géré de manière asynchrone pour le recalcul standard de classeur, c’est pourquoi les tables volumineuses peuvent nécessiter plus de temps de recalcul que le reste du classeur.</span><span class="sxs-lookup"><span data-stu-id="908ff-196">Recalculation is handled asynchronously to regular workbook recalculation so that large tables might take longer to recalculate than the rest of the workbook.</span></span>
    
- <span data-ttu-id="908ff-p117">Les références circulaires sont tolérées. Si le calcul utilisé pour obtenir le résultat dépend d’une ou de plusieurs valeurs de la table de données, Excel ne renvoie pas d’erreur pour la dépendance circulaire.</span><span class="sxs-lookup"><span data-stu-id="908ff-p117">Circular references are tolerated. If the calculation that is used to get the result depends on one or more values from the data table, Excel does not return an error for the circular dependency.</span></span> 

- <span data-ttu-id="908ff-199">Les tables de données n’utilisent pas le calcul à threads multiples.</span><span class="sxs-lookup"><span data-stu-id="908ff-199">Data tables do not use multi-threaded calculation.</span></span>
    
<span data-ttu-id="908ff-p118">Étant donné la manière dont Excel gère le recalcul de tables de données et le fait que les tables volumineuses dépendant de calculs complexes ou longs peuvent nécessiter un certain temps de calcul, Excel vous permet de désactiver le calcul automatique des tables de données. Pour ce faire, définissez le mode de calcul sur Automatique sauf dans les tables. Lorsque le calcul est dans ce mode, l’utilisateur recalcule les tables de données en appuyant sur F9 ou en réalisant une opération de programmation équivalente.</span><span class="sxs-lookup"><span data-stu-id="908ff-p118">Given the different way that Excel handles recalculation of data tables, and the fact that large tables that depend on complex or lengthy calculations can take a long time to calculate, Excel lets you disable the automatic calculation of data tables. To do this, set the calculation mode to Automatic except Data Tables. When calculation is in this mode, the user recalculates the data tables by pressing F9 or some equivalent programmatic operation.</span></span>
  
<span data-ttu-id="908ff-p119">Excel propose des méthodes via lesquelles vous pouvez modifier le mode de recalcul et contrôler ce dernier. Ces méthodes ont été améliorées dans chaque version pour permettre un contrôle plus précis. Les fonctionnalités de l’API C à cet égard reflètent celles qui étaient disponibles dans la version 5 d’Excel, et ne fournissent donc pas le contrôle dont vous pouvez bénéficier en utilisant VBA dans les versions plus récentes.</span><span class="sxs-lookup"><span data-stu-id="908ff-p119">Excel exposes methods through which you can alter the recalculation mode and control recalculation. These methods have been improved from version to version to allow for finer control. The capabilities of the C API in this regard reflect those that were available in Excel version 5, and so do not give you the same control that you have using VBA in more recent versions.</span></span> 
  
<span data-ttu-id="908ff-206">Plus fréquemment utilisées lorsqu’Excel est en mode de calcul manuel, ces méthodes permettent le calcul sélectif des classeurs, des feuilles de calcul et des plages, le recalcul complet de tous les classeurs ouverts, et même la régénération de l’arborescence des dépendances et de la chaîne de calcul.</span><span class="sxs-lookup"><span data-stu-id="908ff-206">Most frequently used when Excel is in manual calculation mode, these methods allow selective calculation of workbooks, worksheets, and ranges, complete recalculation of all open workbooks, and even complete rebuild of the dependency tree and calculation chain.</span></span>
  
### <a name="range-calculation"></a><span data-ttu-id="908ff-207">Calcul de la plage</span><span class="sxs-lookup"><span data-stu-id="908ff-207">Range Calculation</span></span>

<span data-ttu-id="908ff-208">Combinaison de touches : aucune</span><span class="sxs-lookup"><span data-stu-id="908ff-208">Keystroke: None</span></span>
  
<span data-ttu-id="908ff-209">VBA : **Range.Calculate** (introduit dans Excel 2000, modifié dans Excel 2007) et **Range.CalculateRowMajorOrder** (introduit dans Excel 2007)</span><span class="sxs-lookup"><span data-stu-id="908ff-209">VBA: **Range.Calculate** (introduced in Excel 2000, changed in Excel 2007) and **Range.CalculateRowMajorOrder** (introduced in Excel 2007)</span></span> 
  
<span data-ttu-id="908ff-210">API C : non prise en charge</span><span class="sxs-lookup"><span data-stu-id="908ff-210">C API: Not supported</span></span>
  
- <span data-ttu-id="908ff-211">**Mode Manuel**</span><span class="sxs-lookup"><span data-stu-id="908ff-211">**Manual mode**</span></span>
    
    <span data-ttu-id="908ff-p120">Permet de recalculer uniquement les cellules dans la plage donnée, qu’elles soient incorrectes ou non. Le comportement de la méthode **Range.Calculate** est modifié dans Excel 2007. Toutefois, l’ancien comportement est toujours pris en charge par la méthode **Range.CalculateRowMajorOrder**.</span><span class="sxs-lookup"><span data-stu-id="908ff-p120">Recalculates just the cells in the given range regardless of whether they are dirty or not. Behavior of the **Range.Calculate** method changed in Excel 2007; however, the old behavior is still supported by the **Range.CalculateRowMajorOrder** method.</span></span> 
    
- <span data-ttu-id="908ff-214">**Modes Automatique ou Automatique sauf dans les tables**</span><span class="sxs-lookup"><span data-stu-id="908ff-214">**Automatic or Automatic Except Tables modes**</span></span>
    
    <span data-ttu-id="908ff-215">Permet de recalculer le classeur, mais n’impose pas le recalcul de la plage ou des cellules de la plage.</span><span class="sxs-lookup"><span data-stu-id="908ff-215">Recalculates the workbook but does not force recalculation of the range or any cells in the range.</span></span>
    
### <a name="active-worksheet-calculation"></a><span data-ttu-id="908ff-216">Calcul de la feuille de calcul active</span><span class="sxs-lookup"><span data-stu-id="908ff-216">Active Worksheet Calculation</span></span>

<span data-ttu-id="908ff-217">Combinaison de touches : MAJ + F9</span><span class="sxs-lookup"><span data-stu-id="908ff-217">Keystroke: SHIFT+F9</span></span>
  
<span data-ttu-id="908ff-218">VBA : **ActiveSheet.Calculate**</span><span class="sxs-lookup"><span data-stu-id="908ff-218">VBA: **ActiveSheet.Calculate**</span></span>
  
<span data-ttu-id="908ff-219">API C : **xlcCalculateDocument**</span><span class="sxs-lookup"><span data-stu-id="908ff-219">C API: **xlcCalculateDocument**</span></span>
  
- <span data-ttu-id="908ff-220">**Tous les modes**</span><span class="sxs-lookup"><span data-stu-id="908ff-220">**All modes**</span></span>
    
    <span data-ttu-id="908ff-221">Permet de recalculer les cellules marquées pour calcul dans la feuille de calcul active uniquement.</span><span class="sxs-lookup"><span data-stu-id="908ff-221">Recalculates the cells marked for calculation in the active worksheet only.</span></span>
    
### <a name="specified-worksheet-calculation"></a><span data-ttu-id="908ff-222">Calcul de la feuille de calcul indiquée</span><span class="sxs-lookup"><span data-stu-id="908ff-222">Specified Worksheet Calculation</span></span>

<span data-ttu-id="908ff-223">Combinaison de touches : aucune</span><span class="sxs-lookup"><span data-stu-id="908ff-223">Keystroke: None</span></span>
  
<span data-ttu-id="908ff-224">VBA : **Worksheets(** référence **).Calculate**</span><span class="sxs-lookup"><span data-stu-id="908ff-224">VBA: **Worksheets(** reference **).Calculate**</span></span>
  
<span data-ttu-id="908ff-225">API C : non prise en charge</span><span class="sxs-lookup"><span data-stu-id="908ff-225">C API: Not supported</span></span>
  
- <span data-ttu-id="908ff-226">**Tous les modes**</span><span class="sxs-lookup"><span data-stu-id="908ff-226">**All modes**</span></span>
    
    <span data-ttu-id="908ff-p121">Permet de recalculer les cellules incorrectes et leurs dépendances au sein de la feuille de calcul indiquée uniquement. La feuille de calcul porte le nom Référence, comme une chaîne ou le numéro d’index dans le classeur approprié.</span><span class="sxs-lookup"><span data-stu-id="908ff-p121">Recalculates the dirty cells and their dependents within the specified worksheet only. Reference is the name of the worksheet as a string or the index number in the relevant workbook.</span></span>
    
    <span data-ttu-id="908ff-p122">Excel 2000 et les versions ultérieures proposent une propriété de feuille de calcul **Boolean**, la propriété **EnableCalculation**. Le fait de définir cette propriété sur **True** alors qu’elle était définie sur **False** rend toutes les cellules du classeur indiqué incorrectes. En mode automatique, cela déclenche également le recalcul de l’ensemble du classeur.</span><span class="sxs-lookup"><span data-stu-id="908ff-p122">Excel 2000 and later versions expose a **Boolean** worksheet property, the **EnableCalculation** property. Setting this to **True** from **False** dirties all cells in the specified worksheet. In automatic modes, this also triggers a recalculation of the whole workbook.</span></span> 
    
    <span data-ttu-id="908ff-232">En mode Manuel, le code suivant génère le recalcul de la feuille active uniquement.</span><span class="sxs-lookup"><span data-stu-id="908ff-232">In manual mode, the following code causes recalculation of the active sheet only.</span></span>
    
  ```vb
  With ActiveSheet
    .EnableCalculation = False
    .EnableCalculation = True
    .Calculate
  End With
  
  ```

### <a name="workbook-tree-rebuild-and-forced-recalculation"></a><span data-ttu-id="908ff-233">Régénération de l’arborescence du classeur et recalcul forcé</span><span class="sxs-lookup"><span data-stu-id="908ff-233">Workbook Tree Rebuild and Forced Recalculation</span></span>

<span data-ttu-id="908ff-234">Combinaison de touches : CTRL + ALT + MAJ + F9 (introduit dans Excel 2002)</span><span class="sxs-lookup"><span data-stu-id="908ff-234">Keystroke: CTRL+ALT+SHIFT+F9 (introduced in Excel 2002)</span></span>
  
<span data-ttu-id="908ff-235">VBA : **Workbooks(** référence **).ForceFullCalculation** (introduit dans Excel 2007)</span><span class="sxs-lookup"><span data-stu-id="908ff-235">VBA: **Workbooks(** reference **).ForceFullCalculation** (introduced in Excel 2007)</span></span> 
  
<span data-ttu-id="908ff-236">API C : non prise en charge</span><span class="sxs-lookup"><span data-stu-id="908ff-236">C API: Not supported</span></span>
  
- <span data-ttu-id="908ff-237">**Tous les modes**</span><span class="sxs-lookup"><span data-stu-id="908ff-237">**All modes**</span></span>
    
    <span data-ttu-id="908ff-238">Entraîne la régénération de l’arborescence des dépendances et de la chaîne de calcul par Excel pour un classeur donné, et force le recalcul de toutes les cellules qui contiennent des formules.</span><span class="sxs-lookup"><span data-stu-id="908ff-238">Causes Excel to rebuild the dependency tree and the calculation chain for a given workbook and forces a recalculation of all cells that contain formulas.</span></span>
    
### <a name="all-open-workbooks"></a><span data-ttu-id="908ff-239">Tous les classeurs ouverts</span><span class="sxs-lookup"><span data-stu-id="908ff-239">All Open Workbooks</span></span>

<span data-ttu-id="908ff-240">Touche : F9</span><span class="sxs-lookup"><span data-stu-id="908ff-240">Keystroke: F9</span></span>
  
<span data-ttu-id="908ff-241">VBA : **Application.Calculate**</span><span class="sxs-lookup"><span data-stu-id="908ff-241">VBA: **Application.Calculate**</span></span>
  
<span data-ttu-id="908ff-242">API C : **xlcCalculateNow**</span><span class="sxs-lookup"><span data-stu-id="908ff-242">C API: **xlcCalculateNow**</span></span>
  
- <span data-ttu-id="908ff-243">**Tous les modes**</span><span class="sxs-lookup"><span data-stu-id="908ff-243">**All modes**</span></span>
    
    <span data-ttu-id="908ff-p123">Permet de recalculer toutes les cellules qu’Excel a marquées comme incorrectes, c’est-à-dire dépendant de données volatiles ou modifiées, ainsi que les cellules marquées comme incorrectes par programmation. Si le mode de calcul est défini sur Automatique sauf dans les tables, cela calcule les tables qui nécessitent une mise à jour, ainsi que toutes les fonctions volatiles et leurs dépendances.</span><span class="sxs-lookup"><span data-stu-id="908ff-p123">Recalculates all cells that Excel has marked as dirty, that is, dependents of volatile or changed data, and cells programmatically marked as dirty. If the calculation mode is Automatic Except Tables, this calculates those tables that require updating and also all volatile functions and their dependents.</span></span>
    
### <a name="all-open-workbooks-tree-rebuild-and-forced-calculation"></a><span data-ttu-id="908ff-246">Régénération d’arborescence de tous les classeurs ouverts et calcul forcé</span><span class="sxs-lookup"><span data-stu-id="908ff-246">All Open Workbooks Tree Rebuild and Forced Calculation</span></span>

<span data-ttu-id="908ff-247">Combinaison de touches : CTRL + ALT + F9</span><span class="sxs-lookup"><span data-stu-id="908ff-247">Keystroke: CTRL+ALT+F9</span></span>
  
<span data-ttu-id="908ff-248">VBA : **Application.CalculateFull**</span><span class="sxs-lookup"><span data-stu-id="908ff-248">VBA: **Application.CalculateFull**</span></span>
  
<span data-ttu-id="908ff-249">API C : non prise en charge</span><span class="sxs-lookup"><span data-stu-id="908ff-249">C API: Not supported</span></span>
  
- <span data-ttu-id="908ff-250">**Tous les modes**</span><span class="sxs-lookup"><span data-stu-id="908ff-250">**All modes**</span></span>
    
    <span data-ttu-id="908ff-p124">Permet de recalculer toutes les cellules dans tous les classeurs ouverts. Si le mode de calcul est Automatique sauf dans les tables, cela force le recalcul des tables.</span><span class="sxs-lookup"><span data-stu-id="908ff-p124">Recalculates all cells in all open workbooks. If the calculation mode is Automatic Except Tables, it forces the tables to be recalculated.</span></span>
    
## <a name="see-also"></a><span data-ttu-id="908ff-253">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="908ff-253">See also</span></span>



[<span data-ttu-id="908ff-254">Recalcul multithread dans Excel</span><span class="sxs-lookup"><span data-stu-id="908ff-254">Multithreaded Recalculation in Excel</span></span>](multithreaded-recalculation-in-excel.md)
  
[<span data-ttu-id="908ff-255">Types de données utilisées par Excel</span><span class="sxs-lookup"><span data-stu-id="908ff-255">Data Types Used by Excel</span></span>](data-types-used-by-excel.md)
  
[<span data-ttu-id="908ff-256">Gestion de la mémoire dans Excel</span><span class="sxs-lookup"><span data-stu-id="908ff-256">Memory Management in Excel</span></span>](memory-management-in-excel.md)
  
[<span data-ttu-id="908ff-257">Concepts de programmation Excel</span><span class="sxs-lookup"><span data-stu-id="908ff-257">Excel Programming Concepts</span></span>](excel-programming-concepts.md)

