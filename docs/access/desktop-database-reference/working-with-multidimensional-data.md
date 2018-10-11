---
title: Utilisation de données multidimensionnelles
TOCTitle: Working with Multidimensional Data
ms:assetid: a0c9ac73-04da-cfdd-8787-15c8a53ff819
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249740(v=office.15)
ms:contentKeyID: 48546717
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 426a317f378862bd780abe8bce75bb7ff62bedad
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471586"
---
# <a name="working-with-multidimensional-data"></a><span data-ttu-id="d04fb-102">Utilisation de données multidimensionnelles</span><span class="sxs-lookup"><span data-stu-id="d04fb-102">Working with Multidimensional Data</span></span>


<span data-ttu-id="d04fb-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="d04fb-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="d04fb-104">Un *ensemble de cellules* est le résultat d’une requête sur des données multidimensionnelles.</span><span class="sxs-lookup"><span data-stu-id="d04fb-104">A *cellset* is the result of a query on multidimensional data.</span></span> <span data-ttu-id="d04fb-105">Il est constitué d'une collection d'axes, quatre au maximum et deux ou trois en règle générale.</span><span class="sxs-lookup"><span data-stu-id="d04fb-105">It consists of a collection of axes, usually no more than four axes and typically only two or three.</span></span> <span data-ttu-id="d04fb-106">Un *axe* est une collection de membres d’une ou plusieurs dimensions, qui permet de rechercher ou de filtrer des valeurs spécifiques dans un cube.</span><span class="sxs-lookup"><span data-stu-id="d04fb-106">An *axis* is a collection of members from one or more dimensions, which is used to locate or filter specific values in a cube.</span></span>

<span data-ttu-id="d04fb-107">Une *position* est un point le long d’un axe.</span><span class="sxs-lookup"><span data-stu-id="d04fb-107">A *position* is a point along an axis.</span></span> <span data-ttu-id="d04fb-108">Pour un axe comprenant une seule dimension, ces positions constituent un sous-ensemble des membres de la dimension.</span><span class="sxs-lookup"><span data-stu-id="d04fb-108">For an axis consisting of a single dimension, these positions are a subset of the dimension members.</span></span> <span data-ttu-id="d04fb-109">Si un axe comprend plusieurs dimensions, chaque position est une entité composée, qui comprend les éléments *n* où *n* est le nombre de dimensions orientées le long de cet axe.</span><span class="sxs-lookup"><span data-stu-id="d04fb-109">If an axis consists of more than one dimension, then each position is a compound entity, which has *n* parts where *n* is the number of dimensions oriented along that axis.</span></span> <span data-ttu-id="d04fb-110">Chaque partie de la position est un membre d'une dimension constitutive.</span><span class="sxs-lookup"><span data-stu-id="d04fb-110">Each part of the position is a member from one constituent dimension.</span></span>

<span data-ttu-id="d04fb-p103">Par exemple, si les dimensions Géographie et Produit d'un cube contenant des données de ventes sont orientées le long de l'axe des x d'un ensemble de cellules, une position le long de cet axe peut contenir les membres « USA » et « Computers ». Dans cet exemple, les membres de chaque dimension doivent être orientés le long de l'axe des x pour y déterminer une position.</span><span class="sxs-lookup"><span data-stu-id="d04fb-p103">For example, if the Geography and Product dimensions from a cube containing sales data are oriented along the x-axis of a cellset, a position along this axis may contain the members "USA" and "Computers." In this example, determining a position along the x-axis requires that members from each dimension are oriented along the axis.</span></span>

<span data-ttu-id="d04fb-113">Une *cellule* est un objet placé à l’intersection des coordonnées des axes.</span><span class="sxs-lookup"><span data-stu-id="d04fb-113">A *cell* is an object positioned at the intersection of axis coordinates.</span></span> <span data-ttu-id="d04fb-114">Chaque cellule renferme plusieurs informations, notamment les données elles-mêmes, une chaîne mise en forme (la forme affichable des données de la cellule) et la valeur ordinale de la cellule.</span><span class="sxs-lookup"><span data-stu-id="d04fb-114">Each cell has multiple pieces of information associated with it, including the data itself, a formatted string (the displayable form of cell data), and the cell ordinal value.</span></span> <span data-ttu-id="d04fb-115">(Chaque cellule est une valeur ordinale unique de l'ensemble de cellules.</span><span class="sxs-lookup"><span data-stu-id="d04fb-115">(Each cell is a unique ordinal value in the cellset.</span></span> <span data-ttu-id="d04fb-116">La valeur ordinale de la première cellule de l'ensemble est zéro, tandis que celle de la cellule la plus à gauche de la deuxième ligne d'un ensemble de cellules contenant huit colonnes serait huit.)</span><span class="sxs-lookup"><span data-stu-id="d04fb-116">The ordinal value of the first cell in the cellset is zero, while the leftmost cell in the second row of a cellset with eight columns would have an ordinal value of eight.)</span></span>

<span data-ttu-id="d04fb-117">Par exemple, un cube contient les six dimensions suivantes (notez que ce schéma de cube est légèrement différent de celui donné dans la rubrique [Présentation des schémas et données multidimensionnels](overview-of-multidimensional-schemas-and-data.md)) :</span><span class="sxs-lookup"><span data-stu-id="d04fb-117">For example, a cube has the following six dimensions (note that this cube schema differs slightly from the example given in [Overview of Multidimensional Schemas and Data](overview-of-multidimensional-schemas-and-data.md)):</span></span>

  - <span data-ttu-id="d04fb-118">Vendeur</span><span class="sxs-lookup"><span data-stu-id="d04fb-118">Salesperson</span></span>

  - <span data-ttu-id="d04fb-119">Géographie (hiérarchie naturelle)  Continents, Pays, États, etc.</span><span class="sxs-lookup"><span data-stu-id="d04fb-119">Geography (natural hierarchy) — Continents, Countries, States, and so on</span></span>

  - <span data-ttu-id="d04fb-120">Trimestres  Trimestres, Mois, Jours</span><span class="sxs-lookup"><span data-stu-id="d04fb-120">Quarters — Quarters, Months, Days</span></span>

  - <span data-ttu-id="d04fb-121">Années</span><span class="sxs-lookup"><span data-stu-id="d04fb-121">Years</span></span>

  - <span data-ttu-id="d04fb-122">Mesures  Ventes, ChangementPourcentage, VentesBudget</span><span class="sxs-lookup"><span data-stu-id="d04fb-122">Measures — Sales, PercentChange, BudgetedSales</span></span>

  - <span data-ttu-id="d04fb-123">Produits</span><span class="sxs-lookup"><span data-stu-id="d04fb-123">Products</span></span>


> [!NOTE]
> <P><span data-ttu-id="d04fb-124">[!REMARQUE] Les valeurs de cellule de l'exemple peuvent être affichées sous forme de paires classées de valeurs ordinales de position sur les axes où le premier chiffre représente la position sur l'axe des x et le second la position sur l'axe des y.</span><span class="sxs-lookup"><span data-stu-id="d04fb-124">The cell values in the example can be viewed as ordered pairs of axis position ordinals where the first digit represents the x-axis position and the second digit the y-axis position.</span></span></P>



<span data-ttu-id="d04fb-125">Cet ensemble de cellules présente les caractéristiques suivantes :</span><span class="sxs-lookup"><span data-stu-id="d04fb-125">The characteristics of this cellset are as follows:</span></span>

  - <span data-ttu-id="d04fb-126">Dimensions des axes : Trimestres, Vendeur, Géographie</span><span class="sxs-lookup"><span data-stu-id="d04fb-126">Axis dimensions: Quarters, Salesperson, Geography</span></span>

  - <span data-ttu-id="d04fb-127">Dimensions de filtrage : Mesures, Années, Produits</span><span class="sxs-lookup"><span data-stu-id="d04fb-127">Filter dimensions: Measures, Years, Products</span></span>

  - <span data-ttu-id="d04fb-128">Deux axes : COLONNE (x ou Axe 0) et LIGNE (y ou Axe 1)</span><span class="sxs-lookup"><span data-stu-id="d04fb-128">Two axes: COLUMN (x, or Axis 0) and ROW (y, or Axis 1)</span></span>

  - <span data-ttu-id="d04fb-129">axe des x : deux dimensions imbriquées, Vendeur et Géographie</span><span class="sxs-lookup"><span data-stu-id="d04fb-129">x-axis: two nested dimensions, Salesperson and Geography</span></span>

  - <span data-ttu-id="d04fb-130">axe des y : dimension Trimestres</span><span class="sxs-lookup"><span data-stu-id="d04fb-130">y-axis: Quarters dimension</span></span>

<span data-ttu-id="d04fb-p105">L'axe des x comprend deux dimensions imbriquées : Vendeur et Géographie. Quatre membres de la dimension Géographie sont sélectionnés : Seattle, Boston, USA-South et Japan. Deux membres de la dimension Vendeur sont sélectionnés : Valentine et Nash. Ceci produit un total de huit positions sur cet axe (8 = 4\*2).</span><span class="sxs-lookup"><span data-stu-id="d04fb-p105">The x-axis has two nested dimensions: Salesperson and Geography. From Geography, four members are selected: Seattle, Boston, USA-South, and Japan. Two members are selected from Salesperson: Valentine and Nash. This yields a total of eight positions on this axis (8 = 4\*2).</span></span>

<span data-ttu-id="d04fb-135">Chaque coordonnée est représentée sous la forme d'une position avec deux membres  une pour la dimension Vendeur et une autre pour la dimension Géographie :</span><span class="sxs-lookup"><span data-stu-id="d04fb-135">Each coordinate is represented as a position with two members — one from the Salesperson dimension and another from the Geography dimension:</span></span>

```vb 
 
(Valentine, Seattle), (Valentine, Boston), (Valentine, USA_North), 
(Valentine, Japan), (Nash, Seattle), (Nash, Boston), (Nash, USA_North), 
(Nash, Japan) 
```

<span data-ttu-id="d04fb-136">L'axe des y ne comprend qu'une seule dimension, regroupant les huit positions suivantes :</span><span class="sxs-lookup"><span data-stu-id="d04fb-136">The y-axis has only one dimension, containing the following eight positions:</span></span>

```vb 
 
Jan, Feb, Mar, Qtr2, Qtr3, Oct, Nov, Dec 
```

<span data-ttu-id="d04fb-137">Les ensembles de cellules, les cellules, les axes et les positions sont tous représentés par leurs objets ADO MD correspondants : [Cellset](cellset-object-ado-md.md), [Cell](cell-object-ado-md.md), [Axis](axis-object-ado-md.md) et [Position](position-object-ado-md.md).</span><span class="sxs-lookup"><span data-stu-id="d04fb-137">Cellsets, cells, axes, and positions are all represented in ADO MD by corresponding objects: [Cellset](cellset-object-ado-md.md), [Cell](cell-object-ado-md.md), [Axis](axis-object-ado-md.md), and [Position](position-object-ado-md.md).</span></span>
