---
title: Recherche d’une expression exacte dans les pièces jointes des éléments d’un dossier
TOCTitle: Search attachments of items in a folder for an exact phrase
ms:assetid: 3202c0c7-ee3d-4396-b3a9-d24990b44829
ms:mtpsurl: https://msdn.microsoft.com/library/Bb609825(v=office.15)
ms:contentKeyID: 55119889
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: 915ddd65e8227f6cc720c4494ed1767deed121b1
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25406154"
---
# <a name="search-attachments-of-items-in-a-folder-for-an-exact-phrase"></a>Recherche d’une expression exacte dans les pièces jointes des éléments d’un dossier

Cet exemple recherche la chaîne de recherche exacte « office » dans les pièces jointes aux éléments figurant dans la Boîte de réception.

## <a name="example"></a>Exemple

Cet exemple de code utilise une syntaxe DASL (DAV Searching and Locating) pour spécifier une requête. Pour construire le filtre, l’exemple de code vérifie d’abord si la recherche instantanée est activée dans la banque par défaut pour déterminer s’il faut utiliser le mot clé **ci\_phrasematch** pour trouver une correspondance exacte avec « office » dans toutes les pièces jointes. L’exemple applique ensuite le filtre à la méthode [GetTable](https://msdn.microsoft.com/library/bb612592\(v=office.15\)) sur la Boîte de réception et obtient les résultats dans un objet [Table](https://msdn.microsoft.com/library/bb652856\(v=office.15\)). L’exemple de code affiche ensuite l’objet de chaque élément renvoyé dans l’objet **Table**.

L’exemple de code spécifie la propriété **Attachments** d’un élément à l’aide d’une représentation d’espace de noms, https://schemas.microsoft.com/mapi/proptag/0x0EA5001E. La syntaxe à employer pour utiliser le mot clé **ci\_phrasematch** est la suivante :

`<PropertySchemaName> ci_phrasematch <ComparisonString>`

Si vous utilisez Visual Studio pour tester cet exemple de code, vous devez d’abord ajouter une référence au composant Bibliothèque d’objets Microsoft Outlook 15.0 et spécifier la variable lorsque vous importez l’espace de noms **Microsoft.Office.Interop.Outlook**. L'instruction **Importer** ou **utilisation** ne doit pas se produire juste avant les fonctions de l'exemple de code, mais doit être ajoutée avant la déclaration publique. Les lignes de code suivantes montrent comment effectuer l’importation et l’affectation dans Visual Basic et dans C\#.

```vb
Imports Outlook = Microsoft.Office.Interop.Outlook
```


```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```


```vb
Private Sub DemoSearchAttachments()
    Dim filter As String
    Const PR_SEARCH_ATTACHMENTS As String = _
        "https://schemas.microsoft.com/mapi/proptag/0x0EA5001E"
    If (Application.Session.DefaultStore.IsInstantSearchEnabled) Then
        filter = "@SQL=" & Chr(34) _
            & PR_SEARCH_ATTACHMENTS & Chr(34) _
            & " ci_phrasematch 'office'"
        Dim table As Outlook.Table = _
            Application.Session.GetDefaultFolder( _
            Outlook.OlDefaultFolders.olFolderInbox).GetTable( _
            filter, Outlook.OlTableContents.olUserItems)
        While Not (table.EndOfTable)
            Dim row As Outlook.Row = table.GetNextRow()
            Debug.WriteLine(row("Subject"))
        End While
    End If
End Sub
```


```csharp
private void DemoSearchAttachments()
{
    string filter;
    const string PR_SEARCH_ATTACHMENTS =
        "https://schemas.microsoft.com/mapi/proptag/0x0EA5001E";
    if (Application.Session.DefaultStore.IsInstantSearchEnabled)
    {
        filter = "@SQL=" + "\""
            + PR_SEARCH_ATTACHMENTS + "\""
            + " ci_phrasematch 'office'";
        Outlook.Table table = Application.Session.GetDefaultFolder(
            Outlook.OlDefaultFolders.olFolderInbox).GetTable(
            filter, Outlook.OlTableContents.olUserItems);
        while (!table.EndOfTable)
        {
            Outlook.Row row = table.GetNextRow();
            Debug.WriteLine(row["Subject"]);
        }
    }
}
```

## <a name="see-also"></a>Voir aussi

- [Rechercher et filtrer](search-and-filter.md)
