---
title: Envoi d’un élément de courrier électronique avec un compte Hotmail
TOCTitle: Send a mail item by using a Hotmail account
ms:assetid: f25853a7-67c0-46a3-a298-5cdf72ebc53f
ms:mtpsurl: https://msdn.microsoft.com/library/Ff184652(v=office.15)
ms:contentKeyID: 55119797
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: f95d5202f17c21e1c4be4545687f2eee96a00da3
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25407302"
---
# <a name="send-a-mail-item-by-using-a-hotmail-account"></a>Envoi d’un élément de courrier électronique avec un compte Hotmail

Cet exemple utilise la propriété [SendUsingAccount](https://msdn.microsoft.com/library/bb623679\(v=office.15\)) pour envoyer un message en utilisant un compte Windows Live Hotmail.

## <a name="example"></a>Exemple

> [!NOTE] 
> L’exemple de code suivant est un extrait de [programmation d’Applications pour Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493).

Un profil définit un ou plusieurs comptes de messagerie et chaque compte est associé à un serveur d’un type spécifique, comme Microsoft Exchange Server ou POP3 (Post Office Protocol 3). Dans la mesure où votre profil peut inclure plusieurs comptes, indiquez le compte à utiliser pour envoyer l’élément, puis obtenez un objet [Account](https://msdn.microsoft.com/library/bb645103\(v=office.15\)) pour le représenter.

Dans l’exemple de code suivant, un message est créé avec un itinéraire joint, puis il est envoyé à l’aide d’un compte Windows Live Hotmail. Le compte de messagerie Hotmail est utilisé comme objet **Account** dans le profil de l’utilisateur. L’exemple de code définit ensuite la propriété SendUsingAccount sur ce compte et appelle la méthode [Send()](https://msdn.microsoft.com/library/bb644139\(v=office.15\)) à partir de l’objet [MailItem](https://msdn.microsoft.com/library/bb643865\(v=office.15\)).

Si vous utilisez Visual Studio pour tester cet exemple de code, vous devez d’abord ajouter une référence au composant Bibliothèque d’objets Microsoft Outlook 15.0 et spécifier la variable lorsque vous importez l’espace de noms **Microsoft.Office.Interop.Outlook**. L'instruction **d’utilisation** ne doit pas se produire juste avant les fonctions de l'exemple de code, mais doit être ajoutée avant la déclaration publique. La ligne de code suivante montre comment effectuer l’importation et la tâche dans C\#.

```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```

```csharp
private void SendUsingAccountExample()
{
    Outlook.MailItem mail = Application.CreateItem(
        Outlook.OlItemType.olMailItem) as Outlook.MailItem;
    mail.Subject = "Our itinerary";
    mail.Attachments.Add(@"c:\travel\itinerary.doc",
        Outlook.OlAttachmentType.olByValue,
        Type.Missing, Type.Missing);
    Outlook.Account account =
        Application.Session.Accounts["Hotmail"];
    mail.SendUsingAccount = account;
    mail.Send();
}
```

## <a name="see-also"></a>Voir aussi

- [Comptes](accounts.md)
