---
title: Affichage des listes d’adresses d’un profil
TOCTitle: Display the address lists for a profile
ms:assetid: ced8230b-110b-4ccb-a699-588798144154
ms:mtpsurl: https://msdn.microsoft.com/library/Ff184643(v=office.15)
ms:contentKeyID: 55119802
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: 9809603a7c54c9b1ce30c956ebef1a8fb9b6c208
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25405944"
---
# <a name="display-the-address-lists-for-a-profile"></a><span data-ttu-id="f444c-102">Affichage des listes d’adresses d’un profil</span><span class="sxs-lookup"><span data-stu-id="f444c-102">Display the address lists for a profile</span></span>

<span data-ttu-id="f444c-103">Cet exemple montre comment afficher les listes d’adresses du profil actuel.</span><span class="sxs-lookup"><span data-stu-id="f444c-103">This example shows how to display the address lists for the current profile.</span></span>

## <a name="example"></a><span data-ttu-id="f444c-104">Exemple</span><span class="sxs-lookup"><span data-stu-id="f444c-104">Example</span></span>

> [!NOTE] 
> <span data-ttu-id="f444c-105">L’exemple de code suivant est un extrait de la [programmation d’Applications pour Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493).</span><span class="sxs-lookup"><span data-stu-id="f444c-105">The following code example is an excerpt from  [Programming Applications for Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493)  , from Microsoft Press (ISBN 9780735622494, copyright Microsoft Press 2007, all rights reserved).</span></span>

<span data-ttu-id="f444c-p101">Le profil actif contient des listes d'adresses représentées par la collection [AddressLists](https://msdn.microsoft.com/library/bb611894\(v=office.15\)) . Pour obtenir une instance de la collection AddressLists, vous devez utiliser la propriété [AddressLists](https://msdn.microsoft.com/library/bb624048\(v=office.15\)) de l'objet [NameSpace](https://msdn.microsoft.com/library/bb645857\(v=office.15\)) .</span><span class="sxs-lookup"><span data-stu-id="f444c-p101">The current profile contains address lists that are represented by the [AddressLists](https://msdn.microsoft.com/library/bb611894\(v=office.15\)) collection. To get an instance of the AddressLists collection, you must use the [AddressLists](https://msdn.microsoft.com/library/bb624048\(v=office.15\)) property of the [NameSpace](https://msdn.microsoft.com/library/bb645857\(v=office.15\)) object.</span></span>

<span data-ttu-id="f444c-108">Dans l’exemple de code suivant, EnumerateAddressLists commence par énumérer chaque objet [AddressList](https://msdn.microsoft.com/library/bb623538\(v=office.15\)) de la collection AddressLists à l’aide d’une instruction foreach.</span><span class="sxs-lookup"><span data-stu-id="f444c-108">In the following code example,   first enumerates each AddressList object in the AddressLists collection by using a   statement.</span></span> <span data-ttu-id="f444c-109">Une chaîne contenant les valeurs des propriétés [Name](https://msdn.microsoft.com/library/bb609849\(v=office.15\)), [ResolutionOrder](https://msdn.microsoft.com/library/bb646853\(v=office.15\)), [IsReadOnly](https://msdn.microsoft.com/library/bb612676\(v=office.15\)) et [IsInitialAddressList](https://msdn.microsoft.com/library/bb646646\(v=office.15\)) est ensuite créée.</span><span class="sxs-lookup"><span data-stu-id="f444c-109">The example then creates a string that contains the values of the [Name](https://msdn.microsoft.com/library/bb609849\(v=office.15\)) , [ResolutionOrder](https://msdn.microsoft.com/library/bb646853\(v=office.15\)) , [IsReadOnly](https://msdn.microsoft.com/library/bb612676\(v=office.15\)) , and [IsInitialAddressList](https://msdn.microsoft.com/library/bb646646\(v=office.15\)) properties.</span></span> <span data-ttu-id="f444c-110">Enfin, EnumerateAddressLists écrit la chaîne dans les écouteurs de suivi de la collection [Listeners](https://msdn.microsoft.com/library/system.diagnostics.debug.listeners.aspx).</span><span class="sxs-lookup"><span data-stu-id="f444c-110">Finally,   writes the string to the trace listeners of the Listeners collection.</span></span> <span data-ttu-id="f444c-111">Cela permet d’afficher les listes d’adresses du profil actuel.</span><span class="sxs-lookup"><span data-stu-id="f444c-111">This displays each address list for the current profile.</span></span>

<span data-ttu-id="f444c-112">Si vous utilisez Visual Studio pour tester cet exemple de code, vous devez d’abord ajouter une référence au composant Bibliothèque d’objets Microsoft Outlook 15.0 et spécifier la variable lorsque vous importez l’espace de noms **Microsoft.Office.Interop.Outlook**.</span><span class="sxs-lookup"><span data-stu-id="f444c-112">If you use Visual Studio to test this code example, you must first add a reference to the Microsoft Outlook 15.0 Object Library component and specify the   variable when you import the Microsoft.Office.Interop.Outlook namespace.</span></span> <span data-ttu-id="f444c-113">L'instruction **d’utilisation** ne doit pas se produire juste avant les fonctions de l'exemple de code, mais doit être ajoutée avant la déclaration publique.</span><span class="sxs-lookup"><span data-stu-id="f444c-113">The using statement must not occur directly before the functions in the code example but must be added before the public   declaration.</span></span> <span data-ttu-id="f444c-114">La ligne de code suivante montre comment effectuer l’importation et la tâche dans C\#.</span><span class="sxs-lookup"><span data-stu-id="f444c-114">The following line of code shows how to do the import and assignment in C#.</span></span>

```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```

```csharp
private void EnumerateAddressLists()
{
    Outlook.AddressLists addrLists =
         Application.Session.AddressLists;
    foreach (Outlook.AddressList addrList in addrLists)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("Display Name: " + addrList.Name);
        sb.AppendLine("Resolution Order: "
            + addrList.ResolutionOrder.ToString());
        sb.AppendLine("Read-only : "
            + addrList.IsReadOnly.ToString());
        sb.AppendLine("Initial Address List: "
            + addrList.IsInitialAddressList.ToString());
        sb.AppendLine("");
        Debug.WriteLine(sb.ToString());
    }
}
```

## <a name="see-also"></a><span data-ttu-id="f444c-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f444c-115">See also</span></span>

- [<span data-ttu-id="f444c-116">Carnet d’adresses</span><span class="sxs-lookup"><span data-stu-id="f444c-116">Address Book</span></span>](address-book.md)
