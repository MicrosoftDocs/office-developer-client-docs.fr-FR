---
title: Informations détaillées sur HelloData (référence de base de données du bureau Access)
TOCTitle: HelloData Details
ms:assetid: db51e15c-1b5b-c64a-2f84-34dd0e78c6cf
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250105(v=office.15)
ms:contentKeyID: 48548103
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 6de174362e014af3e90686e53a563a10e04ec665
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25470506"
---
# <a name="hellodata-details"></a><span data-ttu-id="64c76-102">Informations détaillées sur HelloData</span><span class="sxs-lookup"><span data-stu-id="64c76-102">HelloData Details</span></span>


<span data-ttu-id="64c76-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="64c76-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="64c76-p101">L'application HelloData effectue les opérations de base d'une application ADO standard : extraction, vérification, modification et mise à jour des données. Lorsque vous démarrez l'application, cliquez sur le premier bouton, **Get Data**. Cette action déclenche l'exécution de la sous-routine GetData().</span><span class="sxs-lookup"><span data-stu-id="64c76-p101">The HelloData application steps through the basic operations of a typical ADO application: getting, examining, editing, and updating data. When you start the application, click the first button, **Get Data**. This will run the GetData() subroutine.</span></span>

## <a name="getdata"></a><span data-ttu-id="64c76-107">GetData</span><span class="sxs-lookup"><span data-stu-id="64c76-107">GetData</span></span>

<span data-ttu-id="64c76-108">GetData place une chaîne de connexion valide dans une variable de niveau module *m\_sConnStr*.</span><span class="sxs-lookup"><span data-stu-id="64c76-108">GetData places a valid connection string into a module-level variable, *m\_sConnStr*.</span></span> <span data-ttu-id="64c76-109">Pour plus d'informations sur les chaînes de connexion, consultez la rubrique [Création de la chaîne de connexion](creating-the-connection-string.md).</span><span class="sxs-lookup"><span data-stu-id="64c76-109">For more information about connection strings, see [Creating the Connection String](creating-the-connection-string.md).</span></span>

<span data-ttu-id="64c76-p103">Affectez un gestionnaire d’erreurs utilisant une instruction Visual Basic **SurErreur**  Pour plus d’informations sur la gestion des erreurs dans ADO, consultez le [Chapitre 6 : Gestion des erreurs](chapter-6-error-handling.md). Un nouvel objet **Connection** est créé et la propriété **CursorLocation** a la valeur **adUseClient** car l’exemple HelloData crée un *jeu de données déconnecté*.  Cela signifie qu’une fois les données extraites de la source de données, la connexion physique à la source de données est interrompue mais que vous pouvez néanmoins continuer à travailler avec les données mises en cache localement dans votre objet **Recordset**</span><span class="sxs-lookup"><span data-stu-id="64c76-p103">Assign an error handler using a Visual Basic **OnError** statement. For more information about error handling in ADO, see [Chapter 6: Error Handling](chapter-6-error-handling.md). A new **Connection** object is created, and the **CursorLocation** property is set to **adUseClient** because the HelloData example creates a *disconnected Recordset*. This means that once the data has been fetched from the data source, the physical connection with the data source is broken, but you can still work with the data that is cached locally in your **Recordset** object.</span></span>

<span data-ttu-id="64c76-114">Après ouverture de la connexion, affectez une chaîne SQL à une variable (sSQL).</span><span class="sxs-lookup"><span data-stu-id="64c76-114">After the connection has been opened, assign a SQL string to a variable (sSQL).</span></span> <span data-ttu-id="64c76-115">Puis instanciez un nouvel objet **Recordset** , m\_oRecordset1.</span><span class="sxs-lookup"><span data-stu-id="64c76-115">Then instantiate a new **Recordset** object, m\_oRecordset1 .</span></span> <span data-ttu-id="64c76-116">Dans la ligne suivante de code, ouvrez le **jeu d’enregistrements** existant **connexion**, en passant.</span><span class="sxs-lookup"><span data-stu-id="64c76-116">In the next line of code, open the **Recordset** over the existing **Connection**, passing in .</span></span> <span data-ttu-id="64c76-117">Dans la ligne suivante de code, ouvrez le **jeu d’enregistrements** existant **connexion**, en passant sSQL comme source de l' **objet Recordset**.</span><span class="sxs-lookup"><span data-stu-id="64c76-117">In the next line of code, open the **Recordset** over the existing **Connection**, passing in sSQL as the source of the **Recordset**.</span></span> <span data-ttu-id="64c76-118">Vous aidez ADO à déterminer que la chaîne SQL transmise en tant que source du **Recordset** est une définition textuelle d'une commande en spécifiant **adCmdText** dans l'argument final de la méthode **Open** du **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="64c76-118">You assist ADO in making the determination that the SQL string you have passed as the source for the **Recordset** is a textual definition of a command by passing **adCmdText** in the final argument to the **Recordset** **Open** method.</span></span> <span data-ttu-id="64c76-119">Cette ligne définit aussi les propriétés **LockType** et **CursorType** associées au **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="64c76-119">This line also sets the **LockType** and **CursorType** associated with the **Recordset**.</span></span>

<span data-ttu-id="64c76-120">La ligne de code suivante précise que la propriété **MarshalOptions** est égale à **adMarshalModifiedOnly**.</span><span class="sxs-lookup"><span data-stu-id="64c76-120">The next line of code sets the **MarshalOptions** property equal to **adMarshalModifiedOnly**.</span></span> <span data-ttu-id="64c76-121">**MarshalOptions** indique au niveau intermédiaire (ou serveur Web) quels sont les enregistrements devant être marshalés.</span><span class="sxs-lookup"><span data-stu-id="64c76-121">**MarshalOptions** indicates which records should be marshaled to the middle tier (or Web server).</span></span> <span data-ttu-id="64c76-122">Pour plus d'informations sur le marshaling, voir la documentation COM.</span><span class="sxs-lookup"><span data-stu-id="64c76-122">For more information about marshaling, see the COM documentation.</span></span> <span data-ttu-id="64c76-123">Lorsque vous utilisez **adMarshalModifiedOnly** avec un curseur côté client ([CursorLocation](cursorlocation-property-ado.md) = **adUseClient**), seuls les enregistrements qui ont été modifiés sur le client sont réécrits au niveau intermédiaire.</span><span class="sxs-lookup"><span data-stu-id="64c76-123">When using **adMarshalModifiedOnly** with a client-side cursor ([CursorLocation](cursorlocation-property-ado.md) = **adUseClient**), only records that have been modified on the client are written back to the middle tier.</span></span> <span data-ttu-id="64c76-124">Le fait de définir la propriété **MarshalOptions** sur **adMarshalModifiedOnly** peut améliorer les performances car les lignes à marshaler sont moins nombreuses.</span><span class="sxs-lookup"><span data-stu-id="64c76-124">Setting **MarshalOptions** to **adMarshalModifiedOnly** can improve performance because fewer rows are marshaled.</span></span>

<span data-ttu-id="64c76-p106">Ensuite, déconnectez le **Recordset** en indiquant que sa propriété **ActiveConnection** est égale à **Nothing**. Pour plus d'informations, consultez la rubrique [Déconnexion et reconnexion de l'objet Recordset](disconnecting-and-reconnecting-the-recordset.md) dans le Chapitre 5 : Mise à jour et persistance des données.</span><span class="sxs-lookup"><span data-stu-id="64c76-p106">Next, disconnect the **Recordset** by setting its **ActiveConnection** property equal to **Nothing**. For more information, see [Disconnecting and Reconnecting the Recordset](disconnecting-and-reconnecting-the-recordset.md) in Chapter 5: Updating and Persisting Data.</span></span>

<span data-ttu-id="64c76-127">Fermez la connexion à la source de données et détruisez l'objet **Connection** existant de façon à libérer les ressources qu'il utilise.</span><span class="sxs-lookup"><span data-stu-id="64c76-127">Close the connection to the data source and destroy the existing **Connection** object, thereby releasing the resources it consumed.</span></span>

<span data-ttu-id="64c76-128">La dernière étape consiste à définir le **Recordset** comme **DataSource** de la grille Microsoft liée aux données sur le formulaire de façon à pouvoir afficher facilement les données du **Recordset** sur ce formulaire.</span><span class="sxs-lookup"><span data-stu-id="64c76-128">The final step is to set the **Recordset** as the **DataSource** for the Microsoft DataBound Grid Control on the form so that you can easily display the data from the **Recordset** on the form.</span></span>

<span data-ttu-id="64c76-p107">Cliquez sur le deuxième bouton, **Examine Data**. Cette action déclenche l'exécution de la sous-routine ExamineData.</span><span class="sxs-lookup"><span data-stu-id="64c76-p107">Click the second button, **Examine Data**. This runs the ExamineData subroutine.</span></span>

## <a name="examinedata"></a><span data-ttu-id="64c76-131">La sous-routine ExamineData</span><span class="sxs-lookup"><span data-stu-id="64c76-131">ExamineData</span></span>

<span data-ttu-id="64c76-132">La sous-routine ExamineData utilise plusieurs méthodes et propriétés de l’objet **Recordset** pour afficher des informations sur les données dans le **jeu d’enregistrements**.</span><span class="sxs-lookup"><span data-stu-id="64c76-132">ExamineData uses various methods and properties of the **Recordset** object to display information about the data in the **Recordset**.</span></span> <span data-ttu-id="64c76-133">Il indique le nombre d’enregistrements à l’aide de la propriété **RecordCount** .</span><span class="sxs-lookup"><span data-stu-id="64c76-133">It reports the number of records by using the **RecordCount** property.</span></span> <span data-ttu-id="64c76-134">Il effectue une boucle dans le **Recordset** et imprime la valeur de la propriété **AbsolutePosition** dans la zone de texte sur le formulaire.</span><span class="sxs-lookup"><span data-stu-id="64c76-134">It loops through the **Recordset** and prints the value of the **AbsolutePosition** property in the display text box on the form.</span></span> <span data-ttu-id="64c76-135">Également dans la boucle, la valeur de la propriété **Bookmark** du troisième enregistrement est placée dans une variable variante *vBookmark*, pour une utilisation ultérieure.</span><span class="sxs-lookup"><span data-stu-id="64c76-135">Also while in the loop, the value of the **Bookmark** property for the third record is placed into a variant variable, *vBookmark*, for later use.</span></span>

<span data-ttu-id="64c76-p109">La routine revient directement au troisième enregistrement grâce à la variable bookmark enregistrée précédemment. Elle appelle la sous-routine WalkFields qui exécute une boucle dans la collection **Fields** du **Recordset** et affiche les détails relatifs à chaque objet **Field** de la collection.</span><span class="sxs-lookup"><span data-stu-id="64c76-p109">The routine navigates directly back to the third record using the bookmark variable that it stored earlier. The routine calls the WalkFields subroutine, which loops through the **Fields** collection of the **Recordset** and displays details about each **Field** in the collection.</span></span>

<span data-ttu-id="64c76-p110">Enfin, ExamineData utilise la propriété **Filter** du **Recordset** pour filtrer les enregistrements selon le critère « CategoryId égale 2 ». L'application de ce filtre a un résultat immédiatement visible dans la grille d'affichage du formulaire.</span><span class="sxs-lookup"><span data-stu-id="64c76-p110">Finally, ExamineData uses the **Filter** property of the **Recordset** to screen for only those records with a CategoryId equal to 2. The result of applying this filter is immediately visible in the display grid on the form.</span></span>

<span data-ttu-id="64c76-140">Pour plus d'informations sur la fonctionnalité de la sous-routine ExamineData, consultez le [chapitre 3 : Examen des données](chapter-3-examining-data.md).</span><span class="sxs-lookup"><span data-stu-id="64c76-140">For more information about the functionality shown in the ExamineData subroutine, see [Chapter 3: Examining Data](chapter-3-examining-data.md).</span></span>

<span data-ttu-id="64c76-p111">Cliquez ensuite sur le troisième bouton, **Edit Data**. Cette action déclenche l'exécution de la sous-routine EditData.</span><span class="sxs-lookup"><span data-stu-id="64c76-p111">Next, click the third button, **Edit Data**. This will run the EditData subroutine.</span></span>

## <a name="editdata"></a><span data-ttu-id="64c76-143">EditData</span><span class="sxs-lookup"><span data-stu-id="64c76-143">EditData</span></span>

<span data-ttu-id="64c76-144">Lorsque le code entre la sous-routine EditData, le **Recordset** est toujours filtré sur CategoryId égale à 2, tel que seuls les éléments qui répondent aux critères de filtre sont visibles.</span><span class="sxs-lookup"><span data-stu-id="64c76-144">When the code enters the EditData subroutine, the **Recordset** is still filtered on CategoryId equal to 2, so only those items that meet the filter criteria are visible.</span></span> <span data-ttu-id="64c76-145">Il parcourt tout d’abord le **jeu d’enregistrements** et augmente le prix de chaque élément visible dans le **jeu d’enregistrements** de 10 pour cent.</span><span class="sxs-lookup"><span data-stu-id="64c76-145">It first loops through the **Recordset** and increases the price of each visible item in the **Recordset** by 10 percent.</span></span> <span data-ttu-id="64c76-146">La valeur du champ **prix** est modifiée en définissant la propriété **Value** de ce champ à un nouveau montant valid.</span><span class="sxs-lookup"><span data-stu-id="64c76-146">The value of the **Price** field is changed by setting the **Value** property for that field equal to a new, valid amount.</span></span>

<span data-ttu-id="64c76-p113">N'oubliez pas que le **Recordset** est déconnecté de la source de données. Les modifications apportées à EditData ne sont effectuées que sur la copie des données dans le cache local. Pour plus d'informations, consultez le [Chapitre 4 : Modification des données](chapter-4-editing-data.md).</span><span class="sxs-lookup"><span data-stu-id="64c76-p113">Remember that the **Recordset** is disconnected from the data source. The changes made in EditData are made only to the locally cached copy of the data. For more information, see [Chapter 4: Editing Data](chapter-4-editing-data.md).</span></span>

<span data-ttu-id="64c76-p114">Aucune modification ne sera apportée à la source de données tant que vous n'aurez pas cliqué sur le quatrième bouton, **Update Data**. Cette action déclenche l'exécution de la sous-routine UpdateData.</span><span class="sxs-lookup"><span data-stu-id="64c76-p114">The changes will not be made on the data source until you click the fourth button, **Update Data**. This will run the UpdateData subroutine.</span></span>

## <a name="updatedata"></a><span data-ttu-id="64c76-152">UpdateData</span><span class="sxs-lookup"><span data-stu-id="64c76-152">UpdateData</span></span>

<span data-ttu-id="64c76-153">La sous-routine UpdateData élimine d'abord le filtre appliqué au **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="64c76-153">UpdateData first removes the filter that has been applied to the **Recordset**.</span></span> <span data-ttu-id="64c76-154">Le code supprime et réinitialise comme **source de données** pour Microsoft liée aux données sur le formulaire afin que le **Recordset** non filtré s’affiche dans la grille.</span><span class="sxs-lookup"><span data-stu-id="64c76-154">The code removes and resets as the **DataSource** for the Microsoft Bound DataGrid on the form so that the unfiltered **Recordset** appears in the grid.</span></span>

<span data-ttu-id="64c76-155">Le code vérifie alors si vous pouvez revenir en arrière dans le **Recordset** en utilisant la méthode **Supports** avec l'argument **adMovePrevious**.</span><span class="sxs-lookup"><span data-stu-id="64c76-155">The code then checks to see whether you can move backward in the **Recordset** by using the **Supports** method with the **adMovePrevious** argument.</span></span>

<span data-ttu-id="64c76-p116">La routine passe au premier enregistrement grâce à la méthode **MoveFirst** et affiche les valeurs d'origine et actuelles du champ par le biais des propriétés **OriginalValue** et **Value** de l'objet **Field**. Ces propriétés, ainsi que la propriété **UnderlyingValue** (non utilisée ici), sont présentées au [Chapitre 5 : Mise à jour et persistance des données](chapter-5-updating-and-persisting-data.md).</span><span class="sxs-lookup"><span data-stu-id="64c76-p116">The routine moves to the first record using the **MoveFirst** method and displays the field's original and current values, using the **OriginalValue** and **Value** properties of the **Field** object. These properties, along with the **UnderlyingValue** property (not used here), are discussed in [Chapter 5: Updating and Persisting Data](chapter-5-updating-and-persisting-data.md).</span></span>

<span data-ttu-id="64c76-p117">Un nouvel objet **Connection** est ensuite créé et est utilisé pour rétablir la connexion à la source de données. Vous reconnectez le **Recordset** à la source de données en définissant le nouvel objet **Connexion** comme **ActiveConnection** du **Recordset**. Pour envoyer les mises à jour au serveur, le code appelle **UpdateBatch** sur le **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="64c76-p117">Next, a new **Connection** object is created and used to reestablish a connection to the data source. You reconnect the **Recordset** to the data source by setting the new **Connection** as the **ActiveConnection** for the **Recordset**. To send the updates to the server, the code calls **UpdateBatch** on the **Recordset**.</span></span>

<span data-ttu-id="64c76-161">Si la mise à jour par lot réussit, une variable d’indicateur de module, est défini sur True.</span><span class="sxs-lookup"><span data-stu-id="64c76-161">If the batch update succeeds, a module-level flag variable, , is set to True.</span></span> <span data-ttu-id="64c76-162">Cet indicateur vous rappellera ultérieurement de nettoyer toutes les modifications apportées à la base de données.</span><span class="sxs-lookup"><span data-stu-id="64c76-162">This will remind you later to clean up all changes made to the database.</span></span>

<span data-ttu-id="64c76-p119">Enfin, le code revient au premier enregistrement du **Recordset** et affiche les valeurs d'origine et actuelles. Les valeurs sont identiques après l'appel de la méthode **UpdateBatch**.</span><span class="sxs-lookup"><span data-stu-id="64c76-p119">Finally, the code moves back to the first record in the **Recordset** and displays the original and current values. The values are the same after the call to **UpdateBatch**.</span></span>

<span data-ttu-id="64c76-165">Pour obtenir des informations plus détaillées sur la mise à jour des données, notamment ce qu'il faut faire lorsque les données du serveur sont modifiées alors que votre **Recordset** est déconnecté, consultez le [Chapitre 5 : Mise à jour et persistance des données](chapter-5-updating-and-persisting-data.md).</span><span class="sxs-lookup"><span data-stu-id="64c76-165">For more detailed information about updating data, including what to do when data on the server changes while your **Recordset** is disconnected, see [Chapter 5: Updating and Persisting Data](chapter-5-updating-and-persisting-data.md).</span></span>

## <a name="formunload"></a><span data-ttu-id="64c76-166">Formulaire\_Unload</span><span class="sxs-lookup"><span data-stu-id="64c76-166">Form\_Unload</span></span>

<span data-ttu-id="64c76-167">Le formulaire\_sous-routine Unload est important pour plusieurs raisons.</span><span class="sxs-lookup"><span data-stu-id="64c76-167">The Form\_Unload subroutine is important for several reasons.</span></span> <span data-ttu-id="64c76-168">Tout d’abord, car il s’agit d’un exemple d’application, formulaire\_Unload nettoie les modifications apportées à la base de données avant la fermeture de l’application.</span><span class="sxs-lookup"><span data-stu-id="64c76-168">First, because this is a sample application, Form\_Unload cleans up the changes made to the database before the application exits.</span></span> <span data-ttu-id="64c76-169">Ensuite, le code montre comment une commande peut être exécutée directement à partir d'un objet **Connection** ouvert grâce à la méthode **Execute**.</span><span class="sxs-lookup"><span data-stu-id="64c76-169">Second, the code shows how a command can be executed directly from an open **Connection** object using the **Execute** method.</span></span> <span data-ttu-id="64c76-170">Enfin, il montre un exemple de l’exécution d’une requête ne retournant-ligne (une requête mise à jour) par rapport à la source de données.</span><span class="sxs-lookup"><span data-stu-id="64c76-170">Finally, it shows an example of executing a non-row–returning query (an UPDATE query) against the data source.</span></span>
