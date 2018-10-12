---
title: Afficher les éléments de demande de tâche envoyés à un destinataire
TOCTitle: Display the task request items sent to a recipient
ms:assetid: 167c3d52-67b5-467c-a7b6-e8cc96002b63
ms:mtpsurl: https://msdn.microsoft.com/library/Ff184591(v=office.15)
ms:contentKeyID: 55119930
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: f26e8d1402b75272e79c6244b7bd2875d783c1f1
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25406637"
---
# <a name="display-the-task-request-items-sent-to-a-recipient"></a><span data-ttu-id="318a5-102">Afficher les éléments de demande de tâche envoyés à un destinataire</span><span class="sxs-lookup"><span data-stu-id="318a5-102">Display the task request items sent to a recipient</span></span>

<span data-ttu-id="318a5-103">Cet exemple montre comment afficher tous les éléments de demande de tâche qui se trouvent dans la boîte de réception d’un destinataire.</span><span class="sxs-lookup"><span data-stu-id="318a5-103">This example shows how to display all of the task request items that are in a recipient's Inbox.</span></span>

## <a name="example"></a><span data-ttu-id="318a5-104">Exemple</span><span class="sxs-lookup"><span data-stu-id="318a5-104">Example</span></span>

> [!NOTE] 
> <span data-ttu-id="318a5-105">L’exemple de code suivant est un extrait de [programmation d’Applications pour Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493).</span><span class="sxs-lookup"><span data-stu-id="318a5-105">The following code example is an excerpt from  [Programming Applications for Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493)  , from Microsoft Press (ISBN 9780735622494, copyright Microsoft Press 2007, all rights reserved).</span></span>

<span data-ttu-id="318a5-106">Un objet [TaskRequestItem](https://msdn.microsoft.com/library/bb610737\(v=office.15\)) représente une demande d’affectation d’une tâche à un autre utilisateur.</span><span class="sxs-lookup"><span data-stu-id="318a5-106">A [TaskRequestItem](https://msdn.microsoft.com/library/bb610737\(v=office.15\)) object represents a request to assign a task to another user.</span></span> <span data-ttu-id="318a5-107">L’objet **TaskRequestItem** est créé lorsque l’élément est reçu dans la boîte de réception du destinataire.</span><span class="sxs-lookup"><span data-stu-id="318a5-107">The **TaskRequestItem** is created when the item is received in the recipient's Inbox.</span></span> <span data-ttu-id="318a5-108">Dans l’exemple de code suivant, ShowTaskRequests filtre la boîte de réception d’un destinataire, crée un objet [Table](https://msdn.microsoft.com/library/bb652856\(v=office.15\)), puis insère une ligne pour chaque élément dont la propriété [MessageClass](https://msdn.microsoft.com/library/bb610592\(v=office.15\)) a la valeur **IPM.TaskRequest**.</span><span class="sxs-lookup"><span data-stu-id="318a5-108">In the following code example,   filters through a recipient's Inbox, creates a Table object, and inserts a row for each item for which the value of the MessageClass property equals IPM.TaskRequest.</span></span> <span data-ttu-id="318a5-109">L’objet de chaque tâche dans le dossier Boîte de réception du destinataire est ensuite inscrit dans les écouteurs de suivi de la collection [Listeners](https://msdn.microsoft.com/library/system.diagnostics.debug.listeners.aspx).</span><span class="sxs-lookup"><span data-stu-id="318a5-109">The subject of each task in the recipient's Inbox folder is then written to the trace listeners of the [Listeners](https://msdn.microsoft.com/library/system.diagnostics.debug.listeners.aspx) collection.</span></span>

<span data-ttu-id="318a5-110">Si vous utilisez Visual Studio pour tester cet exemple de code, vous devez d’abord ajouter une référence au composant Bibliothèque d’objets Microsoft Outlook 15.0 et spécifier la variable lorsque vous importez l’espace de noms **Microsoft.Office.Interop.Outlook**.</span><span class="sxs-lookup"><span data-stu-id="318a5-110">If you use Visual Studio to test this code example, you must first add a reference to the Microsoft Outlook 15.0 Object Library component and specify the   variable when you import the Microsoft.Office.Interop.Outlook namespace.</span></span> <span data-ttu-id="318a5-111">L'instruction **d’utilisation** ne doit pas se produire juste avant les fonctions de l'exemple de code, mais doit être ajoutée avant la déclaration publique.</span><span class="sxs-lookup"><span data-stu-id="318a5-111">The using statement must not occur directly before the functions in the code example but must be added before the public   declaration.</span></span> <span data-ttu-id="318a5-112">La ligne de code suivante montre comment effectuer l’importation et la tâche dans C\#.</span><span class="sxs-lookup"><span data-stu-id="318a5-112">The following line of code shows how to do the import and assignment in C#.</span></span>

```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```


```csharp
private void ShowTaskRequests()
{
    string filter = "[MessageClass] = 'IPM.TaskRequest'";
    Outlook.Table table =
        Application.Session.GetDefaultFolder
        (Outlook.OlDefaultFolders.olFolderInbox).GetTable
        (filter, Outlook.OlTableContents.olUserItems);
    while (!table.EndOfTable)
    {
        Outlook.Row nextRow = table.GetNextRow();
        Debug.WriteLine(nextRow["Subject"]);
    }
}
```

## <a name="see-also"></a><span data-ttu-id="318a5-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="318a5-113">See also</span></span>

- [<span data-ttu-id="318a5-114">Tâches</span><span class="sxs-lookup"><span data-stu-id="318a5-114">Tasks</span></span>](tasks.md)
