---
title: Énumérer et ajouter des catégories
TOCTitle: Enumerate and add categories
ms:assetid: 17a94a01-c463-4332-851e-7d280c66d8c2
ms:mtpsurl: https://msdn.microsoft.com/library/Ff424467(v=office.15)
ms:contentKeyID: 55119829
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: 27fc8630c8e2eb0b491dbb5075f4c58aacec7f93
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25406245"
---
# <a name="enumerate-and-add-categories"></a><span data-ttu-id="7b6bf-102">Énumérer et ajouter des catégories</span><span class="sxs-lookup"><span data-stu-id="7b6bf-102">Enumerate and add categories</span></span>

<span data-ttu-id="7b6bf-103">Cet exemple montre comment énumérer les catégories et ajouter une catégorie à la liste de catégorie principale.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-103">This example shows how to enumerate categories and add a category to the main category list.</span></span>

## <a name="example"></a><span data-ttu-id="7b6bf-104">Exemple</span><span class="sxs-lookup"><span data-stu-id="7b6bf-104">Example</span></span>

> [!NOTE] 
> <span data-ttu-id="7b6bf-105">L’exemple de code suivant est un extrait de [Programmation d’Applications pour Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493).</span><span class="sxs-lookup"><span data-stu-id="7b6bf-105">The following code example is an excerpt from  [Programming Applications for Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493)  , from Microsoft Press (ISBN 9780735622494, copyright Microsoft Press 2007, all rights reserved).</span></span>

<span data-ttu-id="7b6bf-106">Le modèle objet Outlook prend en charge des catégories qui simplifient l'organisation des éléments dans la boîte de réception d'un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-106">The Microsoft Outlook object model supports categories that help organize items in a user's Inbox.</span></span> <span data-ttu-id="7b6bf-107">Pour conserver un niveau supérieur d’organisation, vous pouvez procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="7b6bf-107">To maintain a higher level of organization, you can do the following:</span></span>

- <span data-ttu-id="7b6bf-108">Classer les éléments Outlook et les afficher par catégorie.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-108">Categorize Outlook items and display them by category.</span></span>
- <span data-ttu-id="7b6bf-109">Appliquer plusieurs catégories de couleur à un élément Outlook unique.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-109">Apply multiple color categories to a single Outlook item.</span></span>
- <span data-ttu-id="7b6bf-110">Grouper et trier des éléments Outlook par catégorie de couleur.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-110">Group and sort Outlook items by color category.</span></span>
- <span data-ttu-id="7b6bf-111">Affecter des touches de raccourci à chaque catégorie de couleur, permettant ainsi aux utilisateurs de classer plus facilement les éléments.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-111">Assign shortcut keys to each color category, enabling users to more easily categorize items.</span></span>
- <span data-ttu-id="7b6bf-112">Créer, et changer les catégories de couleurs par programme ou par une action de l'utilisateur dans l'interface utilisateur d'Outlook.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-112">Create, delete, and change color categories either programmatically, or by user action within the Outlook user interface.</span></span>

<span data-ttu-id="7b6bf-113">Pour exposer la fonctionnalité des catégories, le modèle objet Outlook fournit un objet [Category](https://msdn.microsoft.com/library/bb623480\(v=office.15\)) qui représente une catégorie couleur définie par l'utilisateur dans la liste principale de catégories.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-113">To expose the functionality of categories, the Outlook object model provides a [Category](https://msdn.microsoft.com/library/bb623480\(v=office.15\)) object that represents a single user-defined color category in the main category list.</span></span> <span data-ttu-id="7b6bf-114">La liste principale de catégories contient les catégories couleur qui sont présentées dans l'interface utilisateur d'Outlook.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-114">The main category list contains color categories that are presented in the Outlook user interface.</span></span> <span data-ttu-id="7b6bf-115">La liste est représentée par la collection [Categories](https://msdn.microsoft.com/library/bb623535\(v=office.15\)) de l'objet [NameSpace](https://msdn.microsoft.com/library/bb645857\(v=office.15\)) .</span><span class="sxs-lookup"><span data-stu-id="7b6bf-115">The list is represented by the [Categories](https://msdn.microsoft.com/library/bb623535\(v=office.15\)) collection of the [NameSpace](https://msdn.microsoft.com/library/bb645857\(v=office.15\)) object.</span></span> <span data-ttu-id="7b6bf-116">Pour créer un objet **Category**, utilisez la méthode [Add(String, Object, Object)](https://msdn.microsoft.com/library/bb623093\(v=office.15\)) de la collection **Categories**.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-116">To create a **Category** object, use the [Add(String, Object, Object)](https://msdn.microsoft.com/library/bb623093\(v=office.15\)) method of the **Categories** collection.</span></span> <span data-ttu-id="7b6bf-117">Lorsque vous créez un objet **Category**, un identificateur global unique (GUID, Globally Unique Identifier) est créé, et cet identificateur ne peut pas être modifié.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-117">When you create a **Category** object, a globally unique identifier (GUID) is created, and this identifier cannot be changed.</span></span> <span data-ttu-id="7b6bf-118">Il est représenté par la propriété [CategoryID](https://msdn.microsoft.com/library/bb647100\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="7b6bf-118">It is represented by the [CategoryID](https://msdn.microsoft.com/library/bb647100\(v=office.15\)) property.</span></span> <span data-ttu-id="7b6bf-119">Cependant, vous pouvez changer le nom, la couleur et la touche de raccourci associés à une catégorie de couleur en définissant les propriétés [Name](https://msdn.microsoft.com/library/bb645577\(v=office.15\)), [Color](https://msdn.microsoft.com/library/bb612316\(v=office.15\)) et [ShortcutKey](https://msdn.microsoft.com/library/bb644944\(v=office.15\)), respectivement, de l'objet **Category**.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-119">You can, however, change the name, color, and shortcut key associated with a color category by setting the [Name](https://msdn.microsoft.com/library/bb645577\(v=office.15\)) , [Color](https://msdn.microsoft.com/library/bb612316\(v=office.15\)) , and [ShortcutKey](https://msdn.microsoft.com/library/bb644944\(v=office.15\)) properties, respectively, of the **Category** object.</span></span> <span data-ttu-id="7b6bf-120">Vous pouvez modifier la propriété **Color** en définissant ou en récupérant sa constante [OlCategoryColor](https://msdn.microsoft.com/library/bb608974\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="7b6bf-120">You can change the **Color** property by setting or getting its [OlCategoryColor](https://msdn.microsoft.com/library/bb608974\(v=office.15\)) constant.</span></span> <span data-ttu-id="7b6bf-121">Pour reproduire la couleur dans un contrôle personnalisé, utilisez les propriétés en lecture seule suivantes de l'objet **Category**:</span><span class="sxs-lookup"><span data-stu-id="7b6bf-121">To reproduce the color in a custom control, use the following read-only properties of the **Category** object:</span></span>

  - <span data-ttu-id="7b6bf-122">[CategoryBorderColor](https://msdn.microsoft.com/library/bb610083\(v=office.15\))</span><span class="sxs-lookup"><span data-stu-id="7b6bf-122">[CategoryBorderColor Property](https://msdn.microsoft.com/library/bb610083\(v=office.15\))</span></span>

  - <span data-ttu-id="7b6bf-123">[CategoryGradientBottomColor](https://msdn.microsoft.com/library/bb647357\(v=office.15\))</span><span class="sxs-lookup"><span data-stu-id="7b6bf-123">[CategoryGradientBottomColor Property](https://msdn.microsoft.com/library/bb647357\(v=office.15\))</span></span>

  - <span data-ttu-id="7b6bf-124">[CategoryGradientTopColor](https://msdn.microsoft.com/library/bb623975\(v=office.15\))</span><span class="sxs-lookup"><span data-stu-id="7b6bf-124">[CategoryGradientTopColor Property](https://msdn.microsoft.com/library/bb623975\(v=office.15\))</span></span>

<span data-ttu-id="7b6bf-125">Ces propriétés renvoient une valeur **OLE\_COLOR** qui dépend de la propriété **Color** de l'objet **Category**.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-125">These properties return an OLE_COLOR value, which is dependent on the Color property of the Category object.</span></span>

<span data-ttu-id="7b6bf-126">Les éléments Outlook sont présentés en fonction du nom de la catégorie.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-126">Outlook items are displayed based on the category name.</span></span> <span data-ttu-id="7b6bf-127">Chaque élément objet possède une propriété **Category** qui stocke une chaîne séparée par des virgules représentant les noms de catégorie.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-127">Each item object has a **Categories** property that stores a comma-delimited string that represents category names.</span></span> <span data-ttu-id="7b6bf-128">(Par exemple, pour l'objet [MailItem](https://msdn.microsoft.com/library/bb643865\(v=office.15\)), il convient d'utiliser la propriété **MailItem**[Categories](https://msdn.microsoft.com/library/bb646442\(v=office.15\))).</span><span class="sxs-lookup"><span data-stu-id="7b6bf-128">(For example, for the [MailItem](https://msdn.microsoft.com/library/bb643865\(v=office.15\)) object, you would use the **MailItem**[Categories](https://msdn.microsoft.com/library/bb646442\(v=office.15\)) property).</span></span> <span data-ttu-id="7b6bf-129">Cela vous permet d’ajouter une catégorie à l’élément, même si la catégorie n’apparaît pas dans la liste principale de catégories.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-129">This enables you to add a category to the item, even if the category is not present in the main category list.</span></span>


> [!NOTE]
> <span data-ttu-id="7b6bf-p104">[!REMARQUE] Si la propriété **Categories** d'un élément contient un nom de catégorie qui ne figure pas dans la collection **Categories** de l'objet **NameSpace**, le nom de catégorie associé à cet élément Outlook s'affiche, mais sans couleur associée. La propriété **Categories** sur un objet **Item** ne renvoie pas une collection **Categories**.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-p104">If the **Categories** property of an item contains a category name that is not in the **Categories** collection of the **NameSpace** object, the category name associated with that Outlook item is displayed, but without an associated color. The **Categories** property on an **Item** object does not return a **Categories** collection.</span></span>

<span data-ttu-id="7b6bf-132">Dans l'exemple de code suivant, la première procédure, EnumerateCategories, récupère la liste principale de catégories de l'utilisateur actuel, représentée par la collection **Categories**.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-132">In the following code example, the first procedure,  , gets the current user's main list of categories, represented by the Categories collection.</span></span> <span data-ttu-id="7b6bf-133">Elle énumère ensuite les objets **Category** dans cette collection, puis écrit les propriétés **Name** et **CategoryID** des écouteurs de suivi de la collection [Listeners](https://msdn.microsoft.com/library/system.diagnostics.debug.listeners.aspx).</span><span class="sxs-lookup"><span data-stu-id="7b6bf-133">It then enumerates the **Category** objects in that collection, and writes the **Name** and **CategoryID** properties to the trace listeners of the [Listeners](https://msdn.microsoft.com/library/system.diagnostics.debug.listeners.aspx) collection.</span></span> <span data-ttu-id="7b6bf-134">La deuxième procédure, AddACategory, récupère la liste principale de catégories de l'utilisateur actuel et utilise la méthode CategoryExists pour vérifier si une catégorie nommée « ISV » existe dans la collection.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-134">The second procedure,  , gets the current user's main list of categories and uses the   method to check whether a category named "ISV" exists in the collection.</span></span> <span data-ttu-id="7b6bf-135">Si aucune catégorie n'existe sous le nom « ISV », AddACategory ajoute une catégorie nommée « ISV » à la liste principale de catégories et lui affecte la couleur bleu foncé à l'aide de la méthode **Add** de la collection **Categories**.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-135">If no category with the name "ISV" exists,   adds a category named "ISV" to the main category list and assigns it a dark blue color by using the Add method of the Categories collection.</span></span> <span data-ttu-id="7b6bf-136">Il désigne également CTRL + F11 comme touche de raccourci pour la catégorie.</span><span class="sxs-lookup"><span data-stu-id="7b6bf-136">It also assigns CTRL+F11 as the shortcut key for the category.</span></span>

```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```


```csharp
private void EnumerateCategories()
{
    Outlook.Categories categories =
        Application.Session.Categories;
    foreach (Outlook.Category category in categories)
    {
        Debug.WriteLine(category.Name);
        Debug.WriteLine(category.CategoryID);
    }
}

private void AddACategory()
{
    Outlook.Categories categories =
        Application.Session.Categories;
    if (!CategoryExists("ISV"))
    {
        Outlook.Category category = categories.Add("ISV",
            Outlook.OlCategoryColor.olCategoryColorDarkBlue,
            Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF11);
    }
}

private bool CategoryExists(string categoryName)
{
    try
    {
        Outlook.Category category = 
            Application.Session.Categories[categoryName];
        if(category != null)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    catch { return false; }
}
```

## <a name="see-also"></a><span data-ttu-id="7b6bf-137">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7b6bf-137">See also</span></span>

- [<span data-ttu-id="7b6bf-138">Catégories de couleurs</span><span class="sxs-lookup"><span data-stu-id="7b6bf-138">Color Categories</span></span>](color-categories.md)
