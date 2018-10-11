---
title: Accès aux lignes d'un jeu d'enregistrements hiérarchique
TOCTitle: Accessing Rows in a Hierarchical Recordset
ms:assetid: db59b152-b780-539c-17ef-462e8adfb26e
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250106(v=office.15)
ms:contentKeyID: 48548104
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: a7596341f96c7df1d51e67721aacd15f8b075f25
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472272"
---
# <a name="accessing-rows-in-a-hierarchical-recordset"></a><span data-ttu-id="cf463-102">Accès aux lignes d'un jeu d'enregistrements hiérarchique</span><span class="sxs-lookup"><span data-stu-id="cf463-102">Accessing Rows in a Hierarchical Recordset</span></span>


<span data-ttu-id="cf463-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="cf463-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="cf463-104">L'exemple suivant illustre les étapes nécessaires pour accéder aux lignes d'un objet [Recordset](recordset-object-ado.md) hiérarchique :</span><span class="sxs-lookup"><span data-stu-id="cf463-104">The following example shows the steps necessary to access rows in a hierarchical [Recordset](recordset-object-ado.md):</span></span>

1.  <span data-ttu-id="cf463-105">Les objets **Recordset** provenant des tables authors et titleauthor sont liés par l'ID d'auteur.</span><span class="sxs-lookup"><span data-stu-id="cf463-105">**Recordset** objects from the authors and titleauthor tables are related by author ID.</span></span>

2.  <span data-ttu-id="cf463-106">La boucle externe affiche le prénom et le nom de chaque auteur, ainsi que son état et son identification.</span><span class="sxs-lookup"><span data-stu-id="cf463-106">The outer loop displays each author's first and last name, state, and identification.</span></span>

3.  <span data-ttu-id="cf463-107">L'objet **Recordset** ajouté pour chaque ligne est extrait de la collection **Fields** et affecté à *rstTitleAuthor*.</span><span class="sxs-lookup"><span data-stu-id="cf463-107">The appended **Recordset** for each row is retrieved from the **Fields** collection and assigned to *rstTitleAuthor*.</span></span>

4.  <span data-ttu-id="cf463-108">La boucle interne affiche quatre champs de chaque ligne dans l'objet **Recordset** ajouté.</span><span class="sxs-lookup"><span data-stu-id="cf463-108">The inner loop displays four fields from each row in the appended **Recordset**.</span></span>

<span data-ttu-id="cf463-109">(La propriété [StayInSync](stayinsync-property-ado.md) a la valeur FALSE à des fins d’illustration, pour vous pouvez de voir le chapitre changer de manière explicite dans chaque itération de la boucle externe.</span><span class="sxs-lookup"><span data-stu-id="cf463-109">(The [StayInSync](stayinsync-property-ado.md) property is set to FALSE for purposes of illustration — so you can see the chapter change explicitly in each iteration of the outer loop.</span></span> <span data-ttu-id="cf463-110">Toutefois, l’exemple sera plus efficace si l’affectation à l’étape 3 est placée avant la première ligne à l’étape 2, afin que l’affectation est effectuée qu’une seule fois.</span><span class="sxs-lookup"><span data-stu-id="cf463-110">However, the example will be more efficient if the assignment in step 3 is moved before the first line in step 2, so that the assignment is performed only once.</span></span> <span data-ttu-id="cf463-111">Puis définir la propriété **StayInSync** sur TRUE, afin que *rstTitleAuthor* implicitement et automatiquement devient le chapitre correspondant *effectuée* qu’une nouvelle ligne.)</span><span class="sxs-lookup"><span data-stu-id="cf463-111">Then set the **StayInSync** property to TRUE, so that *rstTitleAuthor* will implicitly and automatically change to the corresponding chapter whenever *rst* moves to a new row.)</span></span>

<span data-ttu-id="cf463-112">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="cf463-112">**Example**</span></span>

```vb 
 
Sub datashape() 
 Dim cnn As New ADODB.Connection 
 Dim rst As New ADODB.Recordset 
 Dim rstTitleAuthor As New ADODB.Recordset 
 
 cnn.Provider = "MSDataShape" 
 cnn.Open "Data Provider=MSDASQL;" & _ 
 "Data Source=SRV;" & _ 
 "User Id=MyUserName;Password=MyPassword;Database=Pubs" 
' STEP 1 
 rst.StayInSync = FALSE 
 rst.Open "SHAPE {select * from authors} " & _ 
 "APPEND ({select * from titleauthor} " & _ 
 "RELATE au_id TO au_id) AS chapTitleAuthor", _ 
 cnn 
' STEP 2 
 While Not rst.EOF 
 Debug.Print rst("au_fname"), rst("au_lname"), _ 
 rst("state"), rst("au_id") 
' STEP 3 
 Set rstTitleAuthor = rst("chapTitleAuthor").Value 
' STEP 4 
 While Not rstTitleAuthor.EOF 
 Debug.Print rstTitleAuthor(0), rstTitleAuthor(1), _ 
 rstTitleAuthor(2), rstTitleAuthor(3) 
 rstTitleAuthor.MoveNext 
 Wend 
 rst.MoveNext 
 Wend 
End Sub 
```
