---
title: Visual Basic (référence de base de données du bureau Access)
TOCTitle: Visual Basic
ms:assetid: 9d153b6c-c860-7350-cb3c-b9bd08f75ba8
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249714(v=office.15)
ms:contentKeyID: 48546616
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: eeac988a85f9ef1551d740940d326b67355c478c
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25472336"
---
# <a name="visual-basic"></a><span data-ttu-id="ecdb0-102">Visual Basic</span><span class="sxs-lookup"><span data-stu-id="ecdb0-102">Visual Basic</span></span>


<span data-ttu-id="ecdb0-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="ecdb0-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="ecdb0-p101">Pour gérer les événements ADO dans Microsoft Visual Basic, vous devez déclarer une variable au niveau du module à l'aide du mot-clé **WithEvents**. La variable ne peut être déclarée qu'au sein d'un module de classe et doit être déclarée au niveau du module. Ce n'est toutefois pas aussi restrictif qu'il n'y paraît car les objets **Form** Visual Basic sont également des classes. La manière la plus simple de gérer les événements ADO est de déclarer une variable avec **WithEvents**. L'exemple suivant gère l'événement **ConnectComplete** pour un objet **Connection**:</span><span class="sxs-lookup"><span data-stu-id="ecdb0-p101">In order to handle ADO events in Microsoft Visual Basic, you must declare a module-level variable using the **WithEvents** keyword. The variable can be declared only as part of a class module and must be declared at the module level. This is not as restrictive as it seems, however, because Visual Basic **Form** objects are also classes. The simplest way to handle ADO events is to declare a variable using **WithEvents**. The following example handles the **ConnectComplete** event for a **Connection** object:</span></span>

```vb 
 
' BeginEventExampleVB02 
Dim WithEvents connEvent As Connection 
Attribute connEvent.VB_VarHelpID = -1 
Dim strMsg As String 
 
Private Sub Form_Load() 
 On Error GoTo ErrHandler: 
 
 Dim strConn As String 
 
 ' Create a new object with event 
 ' handling enabled. 
 strConn = "Provider='sqloledb';" & _ 
 "Data Source='MySqlServer';" & _ 
 "Initial Catalog='Northwind';" & _ 
 "Integrated Security='SSPI';" 
 Set connEvent = New ADODB.Connection 
 connEvent.Open strConn 
 
 Exit Sub 
 
ErrHandler: 
 MsgBox strMsg 
End Sub 
 
Private Sub connEvent_ConnectComplete(ByVal pError As ADODB.Error, _ 
 adStatus As ADODB.EventStatusEnum, _ 
 ByVal pConnection As ADODB.Connection) 
 
 If adStatus = adStatusErrorsOccurred Then 
 If Not pError Is Nothing Then 
 Select Case pError.Number 
 Case adErrOperationCancelled 
 ' The operation was cancelled in the 
 ' Will event. Notify the user and exit. 
 strMsg = "I'm sorry you can't connect right now." & vbCrLf 
 strMsg = strMsg & " Click OK to exit." 
 Unload Me 
 Case Else 
 strMsg = "Error " & Format(pError.Number) & vbCrLf 
 strMsg = strMsg & pError.Description 
 strMsg = strMsg & " Click OK to exit." 
 Unload Me 
 End Select 
 Else 
 strMsg = "Error occured. Click OK to exit." 
 Unload Me 
 End If 
 End If 
 'frmWait.btnOK.Enabled = True 
End Sub 
' EndEventExampleVB02 
```

<span data-ttu-id="ecdb0-109">L'objet **Connection** est déclaré au niveau **Form** avec le mot-clé **WithEvents** pour activer la gestion de l'événement.</span><span class="sxs-lookup"><span data-stu-id="ecdb0-109">The **Connection** object is declared at the **Form** level using the **WithEvents** keyword to enable event handling.</span></span> <span data-ttu-id="ecdb0-110">Le formulaire\_Gestionnaire d’événements Load crée l’objet en affectant un nouvel objet **Connection** à *connEvent* et ouvre la connexion.</span><span class="sxs-lookup"><span data-stu-id="ecdb0-110">The Form\_Load event handler actually creates the object by assigning a new **Connection** object to *connEvent* and then opens the connection.</span></span> <span data-ttu-id="ecdb0-111">Bien sûr, une application réelle serait effectuer un traitement plus sous la forme\_Gestionnaire d’événements Load qu’est indiqué ici.</span><span class="sxs-lookup"><span data-stu-id="ecdb0-111">Of course, a real application would do more processing in the Form\_Load event handler than is shown here.</span></span>
