---
title: Obtenir un dossier à partir de son chemin d’accès
TOCTitle: Get a folder based on its folder path
ms:assetid: 613f2209-667c-48f0-82cf-86e3c9a24cb4
ms:mtpsurl: https://msdn.microsoft.com/library/Ff184612(v=office.15)
ms:contentKeyID: 55119858
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: b418cda3f689182cb7133af22e0f2fdc6ba834c2
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25407141"
---
# <a name="get-a-folder-based-on-its-folder-path"></a><span data-ttu-id="e138b-102">Obtenir un dossier à partir de son chemin d’accès</span><span class="sxs-lookup"><span data-stu-id="e138b-102">Get a folder based on its folder path</span></span>

<span data-ttu-id="e138b-103">Cet exemple illustre un chemin d’accès à un dossier et obtient le dossier associé.</span><span class="sxs-lookup"><span data-stu-id="e138b-103">This example takes a folder path and obtains the associated folder.</span></span>

## <a name="example"></a><span data-ttu-id="e138b-104">Exemple</span><span class="sxs-lookup"><span data-stu-id="e138b-104">Example</span></span>

> [!NOTE] 
> <span data-ttu-id="e138b-105">L’exemple de code suivant est un extrait de [programmation d’Applications pour Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493).</span><span class="sxs-lookup"><span data-stu-id="e138b-105">The following code example is an excerpt from  [Programming Applications for Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493)  , from Microsoft Press (ISBN 9780735622494, copyright Microsoft Press 2007, all rights reserved).</span></span>

<span data-ttu-id="e138b-106">Dans l’exemple de code suivant, la méthode GetKeyContacts utilise la propriété [GetRootFolder()](https://msdn.microsoft.com/library/bb645807\(v=office.15\)) pour obtenir le chemin d’accès du dossier Contacts\\Key Contacts.</span><span class="sxs-lookup"><span data-stu-id="e138b-106">In the following code example, the  [](https://msdn.microsoft.com/library/bb645807\(v=office.15\)) method uses the \\ property to obtain the folder path of the Contacts\Key Contacts folder.</span></span> <span data-ttu-id="e138b-107">La méthode GetFolder est ensuite appelée en utilisant la propriété [FolderPath](https://msdn.microsoft.com/library/bb647409\(v=office.15\)) comme argument.</span><span class="sxs-lookup"><span data-stu-id="e138b-107">The   method is then called by using the FolderPath property as the argument.</span></span> <span data-ttu-id="e138b-108">Si GetFolder renvoie un dossier, un message s’affiche indiquant que le dossier Key Contacts a été trouvé.</span><span class="sxs-lookup"><span data-stu-id="e138b-108">If   returns a folder, a message will appear saying the Key Contacts are found.</span></span> <span data-ttu-id="e138b-109">La méthode GetFolder prend un chemin d’accès à un dossier et renvoie l’objet [Folder](https://msdn.microsoft.com/library/bb645774\(v=office.15\)) correct.</span><span class="sxs-lookup"><span data-stu-id="e138b-109">The   method takes in a path to a folder and returns the correct Folder object.</span></span> <span data-ttu-id="e138b-110">Pour ce faire, la propriété **FolderPath** est tout d'abord divisée en tableau de type string, puis ce tableau est utilisé pour trouver l'objet **Folder** correct en commençant par le haut de la propriété **FolderPath**.</span><span class="sxs-lookup"><span data-stu-id="e138b-110">This is done by first splitting the **FolderPath** property into a string array and then using the array to find the correct **Folder** object starting from the top of the **FolderPath** property.</span></span> <span data-ttu-id="e138b-111">Si le dossier spécifié est introuvable, GetFolder renvoie une référence NULL.</span><span class="sxs-lookup"><span data-stu-id="e138b-111">If the specified folder is not found,   returns a null reference.</span></span>

<span data-ttu-id="e138b-112">Si vous utilisez Visual Studio pour tester cet exemple de code, vous devez d’abord ajouter une référence au composant Bibliothèque d’objets Microsoft Outlook 15.0 et spécifier la variable lorsque vous importez l’espace de noms **Microsoft.Office.Interop.Outlook**.</span><span class="sxs-lookup"><span data-stu-id="e138b-112">If you use Visual Studio to test this code example, you must first add a reference to the Microsoft Outlook 15.0 Object Library component and specify the   variable when you import the Microsoft.Office.Interop.Outlook namespace.</span></span> <span data-ttu-id="e138b-113">L'instruction **d’utilisation** ne doit pas se produire juste avant les fonctions de l'exemple de code, mais doit être ajoutée avant la déclaration publique.</span><span class="sxs-lookup"><span data-stu-id="e138b-113">The using statement must not occur directly before the functions in the code example but must be added before the public   declaration.</span></span> <span data-ttu-id="e138b-114">La ligne de code suivante montre comment effectuer l’importation et la tâche dans C\#.</span><span class="sxs-lookup"><span data-stu-id="e138b-114">The following line of code shows how to do the import and assignment in C#.</span></span>

```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```


```csharp
private void GetKeyContacts()
{
    string folderPath =
        Application.Session.
        DefaultStore.GetRootFolder().FolderPath
        + @"\Contacts\Key Contacts";
    Outlook.Folder folder = GetFolder(folderPath);
    if (folder != null)
    {
        //Work with folder here
        Debug.WriteLine("Found Key Contacts");
    }
}

// Returns Folder object based on folder path
private Outlook.Folder GetFolder(string folderPath)
{
    Outlook.Folder folder;
    string backslash = @"\";
    try
    {
        if (folderPath.StartsWith(@"\\"))
        {
            folderPath = folderPath.Remove(0, 2);
        }
        String[] folders =
            folderPath.Split(backslash.ToCharArray());
        folder =
            Application.Session.Folders[folders[0]]
            as Outlook.Folder;
        if (folder != null)
        {
            for (int i = 1; i <= folders.GetUpperBound(0); i++)
            {
                Outlook.Folders subFolders = folder.Folders;
                folder = subFolders[folders[i]]
                    as Outlook.Folder;
                if (folder == null)
                {
                    return null;
                }
            }
        }
        return folder;
    }
    catch { return null; }
}        
```

## <a name="see-also"></a><span data-ttu-id="e138b-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e138b-115">See also</span></span>

- [<span data-ttu-id="e138b-116">Dossiers</span><span class="sxs-lookup"><span data-stu-id="e138b-116">Folders</span></span>](folders.md)
