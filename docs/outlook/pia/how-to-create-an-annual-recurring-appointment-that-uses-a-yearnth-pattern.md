---
title: Création d’un rendez-vous périodique annuel selon une périodicité toutes les n années
TOCTitle: Create an annual recurring appointment that uses a YearNth pattern
ms:assetid: 5fb2ad0b-248c-417d-8868-52e0550d970f
ms:mtpsurl: https://msdn.microsoft.com/library/Ff184611(v=office.15)
ms:contentKeyID: 55119811
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: 9aa18f0e2f875e86acf6cacbf222577e0f9c93e2
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25407078"
---
# <a name="create-an-annual-recurring-appointment-that-uses-a-yearnth-pattern"></a><span data-ttu-id="8c0de-102">Création d’un rendez-vous périodique annuel selon une périodicité toutes les n années</span><span class="sxs-lookup"><span data-stu-id="8c0de-102">Create an annual recurring appointment that uses a YearNth pattern</span></span>

<span data-ttu-id="8c0de-103">Cet exemple montre comment créer un rendez-vous dont le modèle de périodicité annuelle est un jour spécifique comme le premier lundi de juin.</span><span class="sxs-lookup"><span data-stu-id="8c0de-103">This example shows how to create an appointment for which the annual recurrence pattern is a specific day such as the first Monday in June.</span></span>

## <a name="example"></a><span data-ttu-id="8c0de-104">Exemple</span><span class="sxs-lookup"><span data-stu-id="8c0de-104">Example</span></span>

> [!NOTE] 
> <span data-ttu-id="8c0de-105">L’exemple de code suivant est un extrait de la [programmation d’Applications pour Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493).</span><span class="sxs-lookup"><span data-stu-id="8c0de-105">The following code example is an excerpt from  [Programming Applications for Microsoft Office Outlook 2007](https://www.amazon.com/gp/product/0735622493?ie=UTF8&tag=msmsdn-20&linkCode=as2&camp=1789&creative=9325&creativeASIN=0735622493)  , from Microsoft Press (ISBN 9780735622494, copyright Microsoft Press 2007, all rights reserved).</span></span>

<span data-ttu-id="8c0de-106">Si vous voulez créer un rendez-vous annuel qui se produit un jour spécifique de la semaine d'un mois particulier (par exemple, le premier lundi de juin), vous devez utiliser des périodicités YearNth.</span><span class="sxs-lookup"><span data-stu-id="8c0de-106">If you want to create an annual appointment that recurs on a specific day of the week during a specific month (for example, the first Monday in June), you must use YearNth recurrences.</span></span> <span data-ttu-id="8c0de-107">Pour définir une périodicité YearNth, vous devez d’abord définir la propriété [RecurrenceType](https://msdn.microsoft.com/library/bb623463\(v=office.15\)) de l’objet [RecurrencePattern](https://msdn.microsoft.com/library/bb608903\(v=office.15\)) sur olRecursYearNth.</span><span class="sxs-lookup"><span data-stu-id="8c0de-107">To set a YearNth recurrence, you must first set the RecurrenceType property of the RecurrencePattern object to olRecursYearNth.</span></span> <span data-ttu-id="8c0de-108">Puis, vous définissez la propriété [DayOfWeekMask](https://msdn.microsoft.com/library/bb609163\(v=office.15\)) pour spécifier le jour de la semaine où le rendez-vous doit avoir lieu, et la propriété [Instance](https://msdn.microsoft.com/library/bb645269\(v=office.15\)) pour spécifier le numéro d'occurrence du jour de semaine spécifié (par exemple, le troisième jeudi) du mois spécifié pour la périodicité annuelle.</span><span class="sxs-lookup"><span data-stu-id="8c0de-108">Then set the [DayOfWeekMask](https://msdn.microsoft.com/library/bb609163\(v=office.15\)) property to specify on which day of the week the appointment should recur, and the [Instance](https://msdn.microsoft.com/library/bb645269\(v=office.15\)) property to specify the Nth occurrence of the specified day of the week (for example, the third Tuesday) during a specified month for the yearly pattern.</span></span>

<span data-ttu-id="8c0de-109">Lors de l'utilisation d'éléments de rendez-vous périodique, vous devez libérer les références antérieures, le cas échéant, obtenir les nouvelles références à l'élément de rendez-vous périodique avant d'accéder à cet élément ou de le modifier, et libérer ces références dès que vous avez terminé et enregistré les modifications.</span><span class="sxs-lookup"><span data-stu-id="8c0de-109">When you work with recurring appointment items, you should release any prior references, obtain new references to the recurring appointment item before you access or modify the item, and release these references as soon as you are finished and have saved the changes.</span></span> <span data-ttu-id="8c0de-110">Cette pratique s’applique à l’objet **AppointmentItem** périodique, ainsi qu’à tous les objets [Exception](https://msdn.microsoft.com/library/bb610440\(v=office.15\)) ou [RecurrencePattern](https://msdn.microsoft.com/library/bb608903\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="8c0de-110">This practice applies to the recurring **AppointmentItem** object, and any [Exception](https://msdn.microsoft.com/library/bb610440\(v=office.15\)) or [RecurrencePattern](https://msdn.microsoft.com/library/bb608903\(v=office.15\)) object.</span></span> <span data-ttu-id="8c0de-111">Pour publier une référence dans Visual Basic, définissez cet objet existant sur Rien.</span><span class="sxs-lookup"><span data-stu-id="8c0de-111">To release a reference in Visual Basic, set that existing object to Nothing.</span></span> <span data-ttu-id="8c0de-112">Dans C\#, libérez explicitement la mémoire pour cet objet.</span><span class="sxs-lookup"><span data-stu-id="8c0de-112">In C#, explicitly release the memory for that object.</span></span>

<span data-ttu-id="8c0de-p103">Notez que même après la libération de votre référence et la tentative d'obtention d'une nouvelle référence, s'il y a toujours une référence active (détenue par un autre complément ou Outlook) à l'un des objets ci-dessus, la nouvelle référence pointera toujours vers une copie obsolète de l'objet. Par conséquent, il est important que vous libériez vos références dès que vous avez terminé le rendez-vous périodique.</span><span class="sxs-lookup"><span data-stu-id="8c0de-p103">Note that even after you release your reference and attempt to obtain a new reference, if there is still an active reference (held by another add-in or Outlook) to one of the above objects, your new reference will still point to an out-of-date copy of the object. Therefore, it is important that you release your references as soon as you are finished with the recurring appointment.</span></span>

<span data-ttu-id="8c0de-115">Dans l’exemple de code suivant, RecurringYearNthAppointment crée un rendez-vous ayant un modèle de périodicité YearNth.</span><span class="sxs-lookup"><span data-stu-id="8c0de-115">In the following code example,   creates an appointment that has a YearNth recurrence pattern.</span></span> <span data-ttu-id="8c0de-116">RecurringYearNthAppointment crée d’abord un rendez-vous périodique en créant un objet [AppointmentItem](https://msdn.microsoft.com/library/bb645611\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="8c0de-116"> first creates a recurring appointment by creating an AppointmentItem object.</span></span> <span data-ttu-id="8c0de-117">Ensuite, il récupère le modèle de périodicité du rendez-vous en utilisant la méthode [GetRecurrencePattern()](https://msdn.microsoft.com/library/bb652582\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="8c0de-117">Next, it gets the appointment's recurrence pattern by using the [GetRecurrencePattern()](https://msdn.microsoft.com/library/bb652582\(v=office.15\)) method.</span></span> <span data-ttu-id="8c0de-118">Il définit ensuite les propriétés RecurrencePattern suivantes : RecurrenceType, DayOfWeekMask, [MonthOfYear](https://msdn.microsoft.com/library/bb610515\(v=office.15\)), [Instance](https://msdn.microsoft.com/library/bb645269\(v=office.15\)), [Occurrences](https://msdn.microsoft.com/library/bb611303\(v=office.15\)), [Duration](https://msdn.microsoft.com/library/bb644889\(v=office.15\)), [PatternStartDate](https://msdn.microsoft.com/library/bb624492\(v=office.15\)), [StartTime](https://msdn.microsoft.com/library/bb646324\(v=office.15\)) et [EndTime](https://msdn.microsoft.com/library/bb644544\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="8c0de-118">It then sets the following [RecurrencePattern](https://msdn.microsoft.com/library/bb610515\(v=office.15\)) properties: [RecurrenceType](https://msdn.microsoft.com/library/bb645269\(v=office.15\)), [DayOfWeekMask](https://msdn.microsoft.com/library/bb611303\(v=office.15\)), [MonthOfYear](https://msdn.microsoft.com/library/bb644889\(v=office.15\)) , [Instance](https://msdn.microsoft.com/library/bb624492\(v=office.15\)) , [Occurrences , Duration , PatternStartDate , StartTime](https://msdn.microsoft.com/library/bb646324\(v=office.15\)) , and [EndTime](https://msdn.microsoft.com/library/bb644544\(v=office.15\)) .</span></span> <span data-ttu-id="8c0de-119">La propriété MonthOfYear peut prendre une valeur numérique de 1 à 12, où chaque numéro représente le mois correspondant.</span><span class="sxs-lookup"><span data-stu-id="8c0de-119">The MonthOfYear property can take a numerical value of 1 through 12, where each number represents the corresponding month.</span></span> <span data-ttu-id="8c0de-120">Une fois les propriétés définies, RecurringYearNthAppointment enregistre le rendez-vous, puis l’affiche selon le schéma « A lieu le premier lundi de juin à compter du 01/06/2007 jusqu’au 06/06/2016 de 14:00 à 17:00 ».</span><span class="sxs-lookup"><span data-stu-id="8c0de-120">Once the properties are set,   saves the appointment, and then displays it with the pattern "Occurs the first Monday of June effective 6/1/2007 until 6/6/2016 from 2:00 PM to 5:00 PM."</span></span>

<span data-ttu-id="8c0de-121">Si vous utilisez Visual Studio pour tester cet exemple de code, vous devez d’abord ajouter une référence au composant Bibliothèque d’objets Microsoft Outlook 15.0 et spécifier la variable lorsque vous importez l’espace de noms **Microsoft.Office.Interop.Outlook**.</span><span class="sxs-lookup"><span data-stu-id="8c0de-121">If you use Visual Studio to test this code example, you must first add a reference to the Microsoft Outlook 15.0 Object Library component and specify the   variable when you import the Microsoft.Office.Interop.Outlook namespace.</span></span> <span data-ttu-id="8c0de-122">L'instruction **d’utilisation** ne doit pas se produire juste avant les fonctions de l'exemple de code, mais doit être ajoutée avant la déclaration publique.</span><span class="sxs-lookup"><span data-stu-id="8c0de-122">The using statement must not occur directly before the functions in the code example but must be added before the public   declaration.</span></span> <span data-ttu-id="8c0de-123">La ligne de code suivante montre comment effectuer l’importation et la tâche dans C\#.</span><span class="sxs-lookup"><span data-stu-id="8c0de-123">The following line of code shows how to do the import and assignment in C#.</span></span>

```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```

```csharp
private void RecurringYearNthAppointment()
{
    Outlook.AppointmentItem appt = Application.CreateItem(
        Outlook.OlItemType.olAppointmentItem)
        as Outlook.AppointmentItem;
    appt.Subject = "Recurring YearNth Appointment";
    Outlook.RecurrencePattern pattern = appt.GetRecurrencePattern();
    pattern.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearNth;
    pattern.DayOfWeekMask = Outlook.OlDaysOfWeek.olMonday;
    pattern.MonthOfYear = 6;
    pattern.Instance = 1;
    pattern.Occurrences = 10;
    pattern.Duration = 180;
    pattern.PatternStartDate = DateTime.Parse("6/1/2007");
    pattern.StartTime = DateTime.Parse("2:00:00 PM");
    pattern.EndTime = DateTime.Parse("5:00:00 PM");
    appt.Save();
    appt.Display(false);
}
```

## <a name="see-also"></a><span data-ttu-id="8c0de-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8c0de-124">See also</span></span>

- [<span data-ttu-id="8c0de-125">Rendez-vous</span><span class="sxs-lookup"><span data-stu-id="8c0de-125">Appointments</span></span>](appointments.md)
