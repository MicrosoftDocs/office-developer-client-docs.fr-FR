---
title: Ajouter des options de vote à un élément de courrier
TOCTitle: Add voting options to a mail item
ms:assetid: 0fb209a8-178d-411e-9551-0a72e041fd65
ms:mtpsurl: https://msdn.microsoft.com/library/Ff424466(v=office.15)
ms:contentKeyID: 55119867
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: 6a65bdd5086a10b2d6e9803047555a8b052ff38e
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25407169"
---
# <a name="add-voting-options-to-a-mail-item"></a><span data-ttu-id="9cc72-102">Ajouter des options de vote à un élément de courrier</span><span class="sxs-lookup"><span data-stu-id="9cc72-102">Add voting options to a mail item</span></span>

<span data-ttu-id="9cc72-103">Cet exemple montre comment utiliser la propriété [VotingOptions](https://msdn.microsoft.com/library/bb652695\(v=office.15\)) de l’objet [MailItem](https://msdn.microsoft.com/library/bb643865\(v=office.15\)) pour ajouter des options de vote à un message électronique.</span><span class="sxs-lookup"><span data-stu-id="9cc72-103">This example shows how to use the [VotingOptions](https://msdn.microsoft.com/library/bb652695\(v=office.15\)) property of the [MailItem](https://msdn.microsoft.com/library/bb643865\(v=office.15\)) object to add voting options to an e-mail message.</span></span>

## <a name="example"></a><span data-ttu-id="9cc72-104">Exemple</span><span class="sxs-lookup"><span data-stu-id="9cc72-104">Example</span></span>

> [!NOTE] 
> <span data-ttu-id="9cc72-105">L’exemple de code suivant est un extrait de [programmation d’Applications pour Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493).</span><span class="sxs-lookup"><span data-stu-id="9cc72-105">The following code example is an excerpt from  [Programming Applications for Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493)  , from Microsoft Press (ISBN 9780735622494, copyright Microsoft Press 2007, all rights reserved).</span></span>


<span data-ttu-id="9cc72-p101">Les options de vote dans des messages servent à proposer aux destinataires d'un message une liste de choix et à suivre leurs réponses. Pour créer des options de vote par programme, définissez une chaîne correspondant à une liste de valeurs délimitées par des points-virgules pour la propriété **VotingOptions** d'un objet **MailItem**. Les valeurs de la propriété **VotingOptions** apparaissent sous la commande **Vote** du groupe **Répondre** dans le ruban du message reçu.</span><span class="sxs-lookup"><span data-stu-id="9cc72-p101">Voting options on messages are used to give message recipients a list of choices and to track their responses. To create voting options programmatically, set a string that is a semicolon-delimited list of values for the **VotingOptions** property of a **MailItem** object. The values for the **VotingOptions** property will appear under the **Vote** command in the **Respond** group in the ribbon of the received message.</span></span>

<span data-ttu-id="9cc72-109">Dans l’exemple suivant, OrderPizza crée des options de vote dans un nouveau message électronique.</span><span class="sxs-lookup"><span data-stu-id="9cc72-109">In the following example,   creates voting options in a new mail message.</span></span> <span data-ttu-id="9cc72-110">OrderPizza crée d’abord un objet **MailItem**, puis définit la propriété sur **VotingOptions** sur « Cheese; Mushroom; Sausage; Combo; Veg Combo », puis la propriété [Subject](https://msdn.microsoft.com/library/bb611353\(v=office.15\)) sur « Pizza Order ».</span><span class="sxs-lookup"><span data-stu-id="9cc72-110"> first creates a MailItem, and then sets the VotingOptions property to "Cheese; Mushroom; Sausage; Combo; Veg Combo", and the Subject property to "Pizza Order".</span></span> <span data-ttu-id="9cc72-111">Lorsque le message « Pizza Order » est envoyé, les options de vote apparaissent aux destinataires.</span><span class="sxs-lookup"><span data-stu-id="9cc72-111">When the "Pizza Order" message is sent, the voting options appear to recipients.</span></span> <span data-ttu-id="9cc72-112">Pour chaque réponse reçue, le choix du destinataire sera reporté dans la page **Suivi** du message dans le dossier Éléments envoyés de l’expéditeur.</span><span class="sxs-lookup"><span data-stu-id="9cc72-112">For each response received, the recipient's choice will be tallied on the **Tracking** page of the message in the sender's Sent Items folder.</span></span>

<span data-ttu-id="9cc72-113">Si vous utilisez Visual Studio pour tester cet exemple de code, vous devez d’abord ajouter une référence au composant Bibliothèque d’objets Microsoft Outlook 15.0 et spécifier la variable lorsque vous importez l’espace de noms **Microsoft.Office.Interop.Outlook**.</span><span class="sxs-lookup"><span data-stu-id="9cc72-113">If you use Visual Studio to test this code example, you must first add a reference to the Microsoft Outlook 15.0 Object Library component and specify the   variable when you import the Microsoft.Office.Interop.Outlook namespace.</span></span> <span data-ttu-id="9cc72-114">L'instruction **d’utilisation** ne doit pas se produire juste avant les fonctions de l'exemple de code, mais doit être ajoutée avant la déclaration publique.</span><span class="sxs-lookup"><span data-stu-id="9cc72-114">The using statement must not occur directly before the functions in the code example but must be added before the public   declaration.</span></span> <span data-ttu-id="9cc72-115">La ligne de code suivante montre comment effectuer l’importation et la tâche dans C\#.</span><span class="sxs-lookup"><span data-stu-id="9cc72-115">The following line of code shows how to do the import and assignment in C#.</span></span>

```csharp
using Outlook = Microsoft.Office.Interop.Outlook;

    private void OrderPizza()
    {
        Outlook.MailItem mail = (Outlook.MailItem)Application.CreateItem(
            Outlook.OlItemType.olMailItem);
        mail.VotingOptions = “Cheese; Mushroom; Sausage; Combo; Veg Combo;”
        mail.Subject = “Pizza Order”;
        mail.Display(false);
    }
```

## <a name="see-also"></a><span data-ttu-id="9cc72-116">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9cc72-116">See also</span></span>

- [<span data-ttu-id="9cc72-117">Courrier</span><span class="sxs-lookup"><span data-stu-id="9cc72-117">Mail</span></span>](mail.md)
