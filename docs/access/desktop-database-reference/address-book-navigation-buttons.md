---
title: Boutons de navigation dans le Carnet d'adresses
TOCTitle: Address Book Navigation Buttons
ms:assetid: 4ec32c08-5b35-8dce-23ec-0daa4db551cf
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249253(v=office.15)
ms:contentKeyID: 48544765
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 3d68a15873bc5030fd51e54c2219a0ffd91f080f
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471239"
---
# <a name="address-book-navigation-buttons"></a><span data-ttu-id="4878a-102">Boutons de navigation dans le Carnet d'adresses</span><span class="sxs-lookup"><span data-stu-id="4878a-102">Address Book Navigation Buttons</span></span>


<span data-ttu-id="4878a-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="4878a-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="4878a-p101">Les boutons de navigation de l'application Carnet d'adresses sont disposés au bas de la page Web. Vous pouvez utiliser ces boutons pour parcourir les données dans la grille HTML en sélectionnant la première ou la dernière ligne de données, ou les lignes adjacentes à la sélection actuelle.</span><span class="sxs-lookup"><span data-stu-id="4878a-p101">The Address Book application displays the navigation buttons at the bottom of the Web page. You can use the navigation buttons to navigate through the data in the HTML grid display by selecting either the first or last row of data, or rows adjacent to the current selection.</span></span>

## <a name="navigation-sub-procedures"></a><span data-ttu-id="4878a-106">Sous-procédures de navigation</span><span class="sxs-lookup"><span data-stu-id="4878a-106">Navigation Sub Procedures</span></span>

<span data-ttu-id="4878a-107">L'application Carnet d'adresses intègre plusieurs procédures qui permettent aux utilisateurs de parcourir les données en cliquant sur les boutons **Première**, **Suivante**, **Précédente** et **Dernière**.</span><span class="sxs-lookup"><span data-stu-id="4878a-107">The Address Book application contains several procedures that allow users to click the **First**, **Next**, **Previous**, and **Last** buttons to move around the data.</span></span>

<span data-ttu-id="4878a-108">Par exemple, le **premier** bouton permet d’activer le premier VBScript\_procédure OnClick Sub.</span><span class="sxs-lookup"><span data-stu-id="4878a-108">For example, clicking the **First** button activates the VBScript First\_OnClick Sub procedure.</span></span> <span data-ttu-id="4878a-109">Cette procédure exécute une méthode [MoveFirst](movefirst-movelast-movenext-and-moveprevious-methods-rds.md), qui fait de la première ligne de données la sélection actuelle.</span><span class="sxs-lookup"><span data-stu-id="4878a-109">The procedure executes a [MoveFirst](movefirst-movelast-movenext-and-moveprevious-methods-rds.md) method, which makes the first row of data the current selection.</span></span> <span data-ttu-id="4878a-110">En cliquant sur **le bouton** Active la dernière\_procédure OnClick Sub, qui appelle la méthode [MoveLast](movefirst-movelast-movenext-and-moveprevious-methods-rds.md) , qui fait la dernière ligne de données la sélection actuelle.</span><span class="sxs-lookup"><span data-stu-id="4878a-110">Clicking the **Last** button activates the Last\_OnClick Sub procedure, which invokes the [MoveLast](movefirst-movelast-movenext-and-moveprevious-methods-rds.md) method, making the last row of data the current selection.</span></span> <span data-ttu-id="4878a-111">Les autres boutons de navigation fonctionnent de la même façon.</span><span class="sxs-lookup"><span data-stu-id="4878a-111">The remaining navigation buttons work in a similar fashion.</span></span>

```vb 
 
' Move to the first record in the bound Recordset. 
Sub First_OnClick 
 DC1.Recordset.MoveFirst 
End Sub 
 
' Move to the next record from the current position 
' in the bound Recordset. 
Sub Next_OnClick 
 If Not DC1.Recordset.EOF Then ' Cannot move beyond bottom record. 
 DC1.Recordset.MoveNext 
 Exit Sub 
 End If 
End Sub 
 
' Move to the previous record from the current position in the bound 
' Recordset. 
Sub Prev_OnClick 
 If Not DC1.Recordset.BOF Then ' Cannot move beyond top record. 
 DC1.Recordset.MovePrevious 
 Exit Sub 
 End If 
End Sub 
 
' Move to the last record in the bound Recordset. 
Sub Last_OnClick 
 DC1.Recordset.MoveLast 
End Sub 
```
