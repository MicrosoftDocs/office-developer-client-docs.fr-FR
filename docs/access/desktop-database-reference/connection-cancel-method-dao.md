---
title: Connection.Cancel Method (DAO)
TOCTitle: Cancel Method
ms:assetid: 43ad7b64-823d-3fac-e4d4-5e9514f60011
ms:mtpsurl: https://msdn.microsoft.com/library/Ff192953(v=office.15)
ms:contentKeyID: 48544509
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 565dff4e1592ef431aad0c7cb1daa581520f73ec
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472404"
---
# <a name="connectioncancel-method-dao"></a><span data-ttu-id="73adc-102">Connection.Cancel Method (DAO)</span><span class="sxs-lookup"><span data-stu-id="73adc-102">Connection.Cancel Method (DAO)</span></span>

<span data-ttu-id="73adc-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="73adc-103">**Applies to**: Access 2013 | Office 2013</span></span>

## <a name="syntax"></a><span data-ttu-id="73adc-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="73adc-104">Syntax</span></span>

<span data-ttu-id="73adc-105">*expression* . Annuler</span><span class="sxs-lookup"><span data-stu-id="73adc-105">*expression* .Cancel</span></span>

<span data-ttu-id="73adc-106">*expression* Variable qui représente un objet **Connection** .</span><span class="sxs-lookup"><span data-stu-id="73adc-106">*expression* A variable that represents a **Connection** object.</span></span>

## <a name="remarks"></a><span data-ttu-id="73adc-107">Remarques</span><span class="sxs-lookup"><span data-stu-id="73adc-107">Remarks</span></span>

<span data-ttu-id="73adc-108">Utilisez la méthode **Cancel** pour mettre fin à l’exécution d’un appel de méthode **Execute** ou **OpenConnection** asynchrone (autrement dit, la méthode a été appelée avec l’option dbRunAsync).</span><span class="sxs-lookup"><span data-stu-id="73adc-108">Use the **Cancel** method to terminate execution of an asynchronous **Execute** or **OpenConnection** method call (that is, the method was invoked with the dbRunAsync option).</span></span> <span data-ttu-id="73adc-109">**Annuler** renvoie une erreur d’exécution si dbRunAsync n’a pas été utilisé dans la méthode que vous essayez d’interrompre.</span><span class="sxs-lookup"><span data-stu-id="73adc-109">**Cancel** will return a run-time error if dbRunAsync was not used in the method you're trying to terminate.</span></span>

<span data-ttu-id="73adc-110">Une erreur se produira si, après un appel de la méthode **Cancel**, vous tentez de référencer l'objet qui aurait été créé par une appel de la méthode **OpenConnection** asynchrone (à savoir, l'objet **Connection** à partir duquel vous avez appelé la méthode **Cancel**).</span><span class="sxs-lookup"><span data-stu-id="73adc-110">An error will occur if, following a **Cancel** method call, you try to reference the object that would have been created by an asynchronous **OpenConnection** call (that is, the **Connection** object from which you called the **Cancel** method).</span></span>

## <a name="example"></a><span data-ttu-id="73adc-111">Exemple</span><span class="sxs-lookup"><span data-stu-id="73adc-111">Example</span></span>

<span data-ttu-id="73adc-112">Cet exemple utilise la propriété **StillExecuting** et la méthode **Cancel** pour ouvrir un objet **Connection** en mode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="73adc-112">This example uses the **StillExecuting** property and the **Cancel** method to asynchronously open a **Connection** object.</span></span>

```vb
    Sub CancelConnectionX() 
     
     Dim wrkMain As Workspace 
     Dim conMain As Connection 
     Dim sngTime As Single 
     
     Set wrkMain = CreateWorkspace("ODBCWorkspace", _ 
     "admin", "", dbUseODBC) 
     ' Open the connection asynchronously. 
     
     ' Note: The DSN referenced below must be configured to 
     ' use Microsoft Windows NT Authentication Mode to 
     ' authorize user access to the Microsoft SQL Server. 
     Set conMain = wrkMain.OpenConnection("Publishers", _ 
     dbDriverNoPrompt + dbRunAsync, False, _ 
     "ODBC;DATABASE=pubs;DSN=Publishers") 
     
     sngTime = Timer 
     
     ' Wait five seconds. 
     Do While Timer - sngTime < 5 
     Loop 
     
     ' If the connection has not been made, ask the user 
     ' if she wants to keep waiting. If she does not, cancel 
     ' the connection and exit the procedure. 
     Do While conMain.StillExecuting 
     
     If MsgBox("No connection yet--keep waiting?", _ 
     vbYesNo) = vbNo Then 
     conMain.Cancel 
     MsgBox "Connection cancelled!" 
     wrkMain.Close 
     Exit Sub 
     End If 
     
     Loop 
     
     With conMain 
     ' Use the Connection object conMain. 
     .Close 
     End With 
     
     wrkMain.Close 
     
    End Sub
```