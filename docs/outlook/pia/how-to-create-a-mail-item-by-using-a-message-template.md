---
title: Création d’un élément de courrier électronique avec un modèle de message
TOCTitle: Create a mail item by using a message template
ms:assetid: 7d1ffc3b-d1a7-46d1-adb9-ac41e67f626a
ms:mtpsurl: https://msdn.microsoft.com/library/Bb623026(v=office.15)
ms:contentKeyID: 55119862
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: 074a3eb7b83e4556b41549d18b81c97b96db1e6c
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25407099"
---
# <a name="create-a-mail-item-by-using-a-message-template"></a><span data-ttu-id="81ce4-102">Création d’un élément de courrier électronique avec un modèle de message</span><span class="sxs-lookup"><span data-stu-id="81ce4-102">Create a mail item by using a message template</span></span>

<span data-ttu-id="81ce4-103">Cet exemple crée un élément de courrier à l’aide de la méthode [CreateItemFromTemplate](https://msdn.microsoft.com/library/bb611329\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="81ce4-103">This example creates a mail item by using the [CreateItemFromTemplate](https://msdn.microsoft.com/library/bb611329\(v=office.15\)) method.</span></span>

## <a name="example"></a><span data-ttu-id="81ce4-104">Exemple</span><span class="sxs-lookup"><span data-stu-id="81ce4-104">Example</span></span>

<span data-ttu-id="81ce4-105">Cet exemple de code ouvre le fichier de modèle Ivy.oft, affecte un objet, puis enregistre le modèle dans le dossier Brouillons.</span><span class="sxs-lookup"><span data-stu-id="81ce4-105">This code sample opens the Ivy.oft template file, assigns a subject, and then saves the message to the Drafts folder.</span></span>

<span data-ttu-id="81ce4-p101">La méthode **CreateItemFromTemplate** est utile si un fichier de modèle de formulaire Outlook (.oft) que vous voulez utiliser comme modèle de message est stocké sur disque. Le fichier de modèle peut contenir du texte préformaté, du papier à lettres ou des images à inclure dans le message. Toutefois, si le fichier de modèle contient du code derrière le formulaire, le code du formulaire ne s'exécute pas.</span><span class="sxs-lookup"><span data-stu-id="81ce4-p101">The **CreateItemFromTemplate** method is useful if you have an Outlook form template file (.oft) stored on disk that you want to use as a message template. The template file can contain preformatted text, stationery, or images that you want to include in the message. However, if the template file contains code behind the form, the form code will not run.</span></span>

<span data-ttu-id="81ce4-109">Si vous utilisez Visual Studio pour tester cet exemple de code, vous devez d’abord ajouter une référence au composant Bibliothèque d’objets Microsoft Outlook 15.0 et spécifier la variable lorsque vous importez l’espace de noms **Microsoft.Office.Interop.Outlook**.</span><span class="sxs-lookup"><span data-stu-id="81ce4-109">If you use Visual Studio to test this code example, you must first add a reference to the Microsoft Outlook 15.0 Object Library component and specify the   variable when you import the Microsoft.Office.Interop.Outlook namespace.</span></span> <span data-ttu-id="81ce4-110">L'instruction **Importer** ou **utilisation** ne doit pas se produire juste avant les fonctions de l'exemple de code, mais doit être ajoutée avant la déclaration publique.</span><span class="sxs-lookup"><span data-stu-id="81ce4-110">The Imports or using statement must not occur directly before the functions in the code example but must be added before the public   declaration.</span></span> <span data-ttu-id="81ce4-111">Les lignes de code suivantes montrent comment effectuer l’importation et l’affectation dans Visual Basic et dans C\#.</span><span class="sxs-lookup"><span data-stu-id="81ce4-111">The following lines of code show how to do the import and assignment in Visual Basic and C#.</span></span>

```vb
Imports Outlook = Microsoft.Office.Interop.Outlook
```


```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```


```vb
Private Sub CreateItemFromTemplate()
    Dim folder As Outlook.Folder = _
        CType(Application.Session.GetDefaultFolder( _
        Outlook.OlDefaultFolders.olFolderDrafts), Outlook.Folder)
    Dim mail As Outlook.MailItem = _
        CType(Application.CreateItemFromTemplate( _
        "c:\ivy.oft", folder), Outlook.MailItem)
    mail.Subject = "Congratulations"
    mail.Save()
End Sub
```


```csharp
private void CreateItemFromTemplate()
{
    Outlook.Folder folder =
        Application.Session.GetDefaultFolder(
        Outlook.OlDefaultFolders.olFolderDrafts) as Outlook.Folder;
    Outlook.MailItem mail =
        Application.CreateItemFromTemplate(
        @"c:\ivy.oft", folder) as Outlook.MailItem;
    mail.Subject = "Congratulations";
    mail.Save();
}
```

## <a name="see-also"></a><span data-ttu-id="81ce4-112">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="81ce4-112">See also</span></span>

- [<span data-ttu-id="81ce4-113">Courrier</span><span class="sxs-lookup"><span data-stu-id="81ce4-113">Mail</span></span>](mail.md)
