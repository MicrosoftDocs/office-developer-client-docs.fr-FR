---
title: Obtention des informations de l’utilisateur actuel
TOCTitle: Get information about the current user
ms:assetid: 3802523a-3ccf-4cca-a348-abe2645a0d9c
ms:mtpsurl: https://msdn.microsoft.com/library/Ff184601(v=office.15)
ms:contentKeyID: 55119840
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: 7ac10c01d205f4f4590183bcef8006e8f1344cbd
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25406973"
---
# <a name="get-information-about-the-current-user"></a><span data-ttu-id="408e6-102">Obtention des informations de l’utilisateur actuel</span><span class="sxs-lookup"><span data-stu-id="408e6-102">Get information about the current login user.</span></span>

<span data-ttu-id="408e6-103">Cet exemple montre comment obtenir les informations (par exemple, le nom, la fonction et les numéros de téléphone) de l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="408e6-103">This example shows how to get the current user’s information, such as name, job title, and telephone number.</span></span>

## <a name="example"></a><span data-ttu-id="408e6-104">Exemple</span><span class="sxs-lookup"><span data-stu-id="408e6-104">Example</span></span>

> [!NOTE] 
> <span data-ttu-id="408e6-105">L’exemple de code suivant est un extrait de la [programmation d’Applications pour Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493).</span><span class="sxs-lookup"><span data-stu-id="408e6-105">The following code example is an excerpt from  [Programming Applications for Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493)  , from Microsoft Press (ISBN 9780735622494, copyright Microsoft Press 2007, all rights reserved).</span></span>

<span data-ttu-id="408e6-106">Pour obtenir un objet [ExchangeUser](https://msdn.microsoft.com/library/bb609574\(v=office.15\)) à partir d’un objet [AddressEntry](https://msdn.microsoft.com/library/bb609728\(v=office.15\)), appelez la méthode [GetExchangeUser()](https://msdn.microsoft.com/library/bb611808\(v=office.15\)) sur l’objet **AddressEntry**.</span><span class="sxs-lookup"><span data-stu-id="408e6-106">To obtain an [ExchangeUser](https://msdn.microsoft.com/library/bb609574\(v=office.15\)) object from an [AddressEntry](https://msdn.microsoft.com/library/bb609728\(v=office.15\)) object, call the [GetExchangeUser()](https://msdn.microsoft.com/library/bb611808\(v=office.15\)) method on the **AddressEntry** object.</span></span> <span data-ttu-id="408e6-107">Dans la procédure suivante, GetCurrentUserInfo obtient la propriété [AddressEntry](https://msdn.microsoft.com/library/bb644359\(v=office.15\)) de l’objet [Recipient](https://msdn.microsoft.com/library/bb624370\(v=office.15\)) en utilisant la propriété [CurrentUser](https://msdn.microsoft.com/library/bb622574\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="408e6-107">In the following procedure,   gets the AddressEntry property for the Recipient object by using the CurrentUser property.</span></span> <span data-ttu-id="408e6-108">Si l’objet **AddressEntry** représente un utilisateur de boîte aux lettres Exchange, GetCurrentUserInfo appelle la méthode **GetExchangeUser** et un objet **ExchangeUser** est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="408e6-108">If the AddressEntry object represents an Exchange mailbox user,   calls the GetExchangeUser method and an ExchangeUser object is returned.</span></span> <span data-ttu-id="408e6-109">Les propriétés [Name](https://msdn.microsoft.com/library/bb622941\(v=office.15\)), [PrimarySmtpAddress](https://msdn.microsoft.com/library/bb645506\(v=office.15\)), [JobTitle](https://msdn.microsoft.com/library/bb645451\(v=office.15\)), [Department](https://msdn.microsoft.com/library/bb623789\(v=office.15\)), [OfficeLocation](https://msdn.microsoft.com/library/bb611429\(v=office.15\)), [BusinessTelephoneNumber](https://msdn.microsoft.com/library/bb612294\(v=office.15\)) et [MobileTelephoneNumber](https://msdn.microsoft.com/library/bb609292\(v=office.15\)) sont écrites dans les écouteurs de suivi de la collection [Listeners](https://msdn.microsoft.com/library/system.diagnostics.debug.listeners.aspx).</span><span class="sxs-lookup"><span data-stu-id="408e6-109">The [Name](https://msdn.microsoft.com/library/bb622941\(v=office.15\)) , [PrimarySmtpAddress](https://msdn.microsoft.com/library/bb645506\(v=office.15\)) , [JobTitle](https://msdn.microsoft.com/library/bb645451\(v=office.15\)) , [Department](https://msdn.microsoft.com/library/bb623789\(v=office.15\)) , [OfficeLocation](https://msdn.microsoft.com/library/bb611429\(v=office.15\)) , [BusinessTelephoneNumber](https://msdn.microsoft.com/library/bb612294\(v=office.15\)) , and [MobileTelephoneNumber](https://msdn.microsoft.com/library/bb609292\(v=office.15\)) properties are written to the trace listeners of the [Listeners](https://msdn.microsoft.com/library/system.diagnostics.debug.listeners.aspx) collection.</span></span>

<span data-ttu-id="408e6-110">Si vous utilisez Visual Studio pour tester cet exemple de code, vous devez d’abord ajouter une référence au composant Bibliothèque d’objets Microsoft Outlook 15.0 et spécifier la variable lorsque vous importez l’espace de noms **Microsoft.Office.Interop.Outlook**.</span><span class="sxs-lookup"><span data-stu-id="408e6-110">If you use Visual Studio to test this code example, you must first add a reference to the Microsoft Outlook 15.0 Object Library component and specify the   variable when you import the Microsoft.Office.Interop.Outlook namespace.</span></span> <span data-ttu-id="408e6-111">L'instruction **d’utilisation** ne doit pas se produire juste avant les fonctions de l'exemple de code, mais doit être ajoutée avant la déclaration publique.</span><span class="sxs-lookup"><span data-stu-id="408e6-111">The using statement must not occur directly before the functions in the code example but must be added before the public   declaration.</span></span> <span data-ttu-id="408e6-112">La ligne de code suivante montre comment effectuer l’importation et la tâche dans C\#.</span><span class="sxs-lookup"><span data-stu-id="408e6-112">The following line of code shows how to do the import and assignment in C#.</span></span>

```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```


```csharp
private void GetCurrentUserInfo()
{
    Outlook.AddressEntry addrEntry =
        Application.Session.CurrentUser.AddressEntry;
    if (addrEntry.Type == "EX")
    {
        Outlook.ExchangeUser currentUser =
            Application.Session.CurrentUser.
            AddressEntry.GetExchangeUser();
        if (currentUser != null)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Name: "
                + currentUser.Name);
            sb.AppendLine("STMP address: "
                + currentUser.PrimarySmtpAddress);
            sb.AppendLine("Title: "
                + currentUser.JobTitle);
            sb.AppendLine("Department: "
                + currentUser.Department);
            sb.AppendLine("Location: "
                + currentUser.OfficeLocation);
            sb.AppendLine("Business phone: "
                + currentUser.BusinessTelephoneNumber);
            sb.AppendLine("Mobile phone: "
                + currentUser.MobileTelephoneNumber);
            Debug.WriteLine(sb.ToString());
        }
    }
}
```

## <a name="see-also"></a><span data-ttu-id="408e6-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="408e6-113">See also</span></span>

- [<span data-ttu-id="408e6-114">Utilisateurs Exchange</span><span class="sxs-lookup"><span data-stu-id="408e6-114">Exchange Users</span></span>](exchange-users.md)
