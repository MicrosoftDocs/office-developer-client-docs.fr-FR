---
title: Workspace.CommitTrans Method (DAO)
TOCTitle: CommitTrans Method
ms:assetid: e6d129fb-a578-5c79-9c16-6444519f0daf
ms:mtpsurl: https://msdn.microsoft.com/library/Ff835985(v=office.15)
ms:contentKeyID: 48548391
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 149727c7f7424a7508edb21bf486e6148cd12295
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25469292"
---
# <a name="workspacecommittrans-method-dao"></a><span data-ttu-id="1954d-102">Workspace.CommitTrans Method (DAO)</span><span class="sxs-lookup"><span data-stu-id="1954d-102">Workspace.CommitTrans Method (DAO)</span></span>

<span data-ttu-id="1954d-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="1954d-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="1954d-104">Met fin à la transaction en cours et enregistre les modifications.</span><span class="sxs-lookup"><span data-stu-id="1954d-104">Ends the current transaction and saves the changes.</span></span>

## <a name="syntax"></a><span data-ttu-id="1954d-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="1954d-105">Syntax</span></span>

<span data-ttu-id="1954d-106">*expression* . CommitTrans (***Options***)</span><span class="sxs-lookup"><span data-stu-id="1954d-106">*expression* .CommitTrans(***Options***)</span></span>

<span data-ttu-id="1954d-107">*expression* Variable qui représente un objet **Workspace** .</span><span class="sxs-lookup"><span data-stu-id="1954d-107">*expression* A variable that represents a **Workspace** object.</span></span>

### <a name="parameters"></a><span data-ttu-id="1954d-108">Paramètres</span><span class="sxs-lookup"><span data-stu-id="1954d-108">Parameters</span></span>

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th><p><span data-ttu-id="1954d-109">Name</span><span class="sxs-lookup"><span data-stu-id="1954d-109">Name</span></span></p></th>
<th><p><span data-ttu-id="1954d-110">Obligatoire/Facultatif</span><span class="sxs-lookup"><span data-stu-id="1954d-110">Required/Optional</span></span></p></th>
<th><p><span data-ttu-id="1954d-111">Type de données</span><span class="sxs-lookup"><span data-stu-id="1954d-111">Data Type</span></span></p></th>
<th><p><span data-ttu-id="1954d-112">Description</span><span class="sxs-lookup"><span data-stu-id="1954d-112">Description</span></span></p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><span data-ttu-id="1954d-113">Option</span><span class="sxs-lookup"><span data-stu-id="1954d-113">Option</span></span></p></td>
<td><p><span data-ttu-id="1954d-114">Facultatif</span><span class="sxs-lookup"><span data-stu-id="1954d-114">Optional</span></span></p></td>
<td><p><span data-ttu-id="1954d-115"><strong>Entier long</strong></span><span class="sxs-lookup"><span data-stu-id="1954d-115"><strong>Long</strong></span></span></p></td>
<td><p><span data-ttu-id="1954d-p101">Dans un espace de travail Microsoft Access, vous pouvez inclure la constante <strong>dbForceOSFlush</strong> avec <strong>CommitTrans</strong>. Cela force le moteur de base de données à purger toutes les mises à jour sur le disque au lieu de les mettre en mémoire cache temporairement. Sans cette option, un utilisateur peut reprendre le contrôle dès que le programme d’application appelle <strong>CommitTrans</strong>, éteindre l’ordinateur et les données ne sont pas écrites sur le disque. Même si cette option peut affecter les performances de votre application, elle s’avère utile dans les situations dans lesquelles l’ordinateur a pu être fermé avant que les mises à jour mises en mémoire cache ne soient enregistrées sur le disque.</span><span class="sxs-lookup"><span data-stu-id="1954d-p101">In a Microsoft Access workspace, you can include the <strong>dbForceOSFlush</strong> constant with <strong>CommitTrans</strong>. This forces the database engine to immediately flush all updates to disk, instead of caching them temporarily. Without using this option, a user could get control back immediately after the application program calls <strong>CommitTrans</strong>, turn the computer off, and not have the data written to disk. While using this option may affect your application's performance, it is useful in situations where the computer could be shut off before cached updates are saved to disk.</span></span></p></td>
</tr>
</tbody>
</table>


## <a name="remarks"></a><span data-ttu-id="1954d-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="1954d-120">Remarks</span></span>

<span data-ttu-id="1954d-p102">Les méthodes de transaction **BeginTrans**, **CommitTrans** et **Rollback** gèrent le traitement des transactions au cours d'une session définie par un objet **Workspace**. Vous utilisez ces méthodes avec un objet **Workspace** lorsque vous souhaitez traiter une série de modifications apportées en bloc aux bases de données dans une session.</span><span class="sxs-lookup"><span data-stu-id="1954d-p102">The transaction methods **BeginTrans**, **CommitTrans**, and **Rollback** manage transaction processing during a session defined by a **Workspace** object. You use these methods with a **Workspace** object when you want to treat a series of changes made to the databases in a session as one unit.</span></span>

<span data-ttu-id="1954d-p103">En général, vous utilisez des transactions pour conserver l'intégrité de vos données lorsque vous devez mettre à jour des enregistrements dans deux ou plusieurs tables et vous assurer que les modifications sont terminées (validées) dans toutes les tables ou dans aucune des tables (annulées). Par exemple, si vous transférez de l'argent d'un compte à un autre, vous devez soustraire un montant d'un compte et l'ajouter à un autre compte. Si une de ces mises à jour échoue, les comptes ne sont plus équilibrés. Utilisez la méthode **BeginTrans** avant de mettre à jour le premier enregistrement, puis, si une mise à jour suivante échoue, vous pouvez utiliser la méthode **Rollback** pour annuler toutes les mises à jour. Utilisez la méthode **CommitTrans** une fois que vous avez réussi à mettre à jour le dernier enregistrement.</span><span class="sxs-lookup"><span data-stu-id="1954d-p103">Typically, you use transactions to maintain the integrity of your data when you must both update records in two or more tables and ensure changes are completed (committed) in all tables or none at all (rolled back). For example, if you transfer money from one account to another, you might subtract an amount from one and add the amount to another. If either update fails, the accounts no longer balance. Use the **BeginTrans** method before updating the first record, and then, if any subsequent update fails, you can use the **Rollback** method to undo all of the updates. Use the **CommitTrans** method after you successfully update the last record.</span></span>

> [!NOTE]
> <span data-ttu-id="1954d-p104">[!REMARQUE] Dans un objet **Workspace**, les transactions sont toujours globales dans l'objet **Workspace** et vous n'êtes pas limité à un seul objet **Connection** ou **Database**. Si vous effectuez des opérations sur plusieurs connexions ou bases de données dans une transaction **Workspace**, la résolution de la transaction (à l'aide de la méthode **CommitTrans** ou **Rollback**) affecte toutes les opérations de toutes les connexions et bases de données dans cet espace de travail.</span><span class="sxs-lookup"><span data-stu-id="1954d-p104">Within one **Workspace** object, transactions are always global to the **Workspace** and aren't limited to only one **Connection** or **Database** object. If you perform operations on more than one connection or database within a **Workspace** transaction, resolving the transaction (that is, using the **CommitTrans** or **Rollback** method) affects all operations on all connections and databases within that workspace.</span></span>

<span data-ttu-id="1954d-p105">Une fois que vous avez utilisé la méthode **CommitTrans**, vous ne pouvez pas annuler les modifications effectuées au cours de cette transaction sauf si la transaction est imbriquée dans une autre transaction qui est elle-même annulée. Si vous imbriquez des transactions, vous devez résoudre la transaction active avant de pouvoir résoudre une transaction située à un niveau supérieur d'imbrication.</span><span class="sxs-lookup"><span data-stu-id="1954d-p105">After you use **CommitTrans**, you can't undo changes made during that transaction unless the transaction is nested within another transaction that is itself rolled back. If you nest transactions, you must resolve the current transaction before you can resolve a transaction at a higher level of nesting.</span></span>

<span data-ttu-id="1954d-132">Si vous souhaitez utiliser des transactions simultanées avec des portées non imbriquées qui se chevauchent, vous pouvez créer d'autres objets **Workspace** qui contiennent les transactions concomitantes.</span><span class="sxs-lookup"><span data-stu-id="1954d-132">If you want to have simultaneous transactions with overlapping, non-nested scopes, you can create additional **Workspace** objects to contain the concurrent transactions.</span></span>

<span data-ttu-id="1954d-133">Si vous fermez un objet **Workspace** sans résoudre les transactions en attente, les transactions sont automatiquement annulées.</span><span class="sxs-lookup"><span data-stu-id="1954d-133">If you close a **Workspace** object without resolving any pending transactions, the transactions are automatically rolled back.</span></span>

<span data-ttu-id="1954d-134">Si vous utilisez la méthode **CommitTrans** ou **Rollback** sans utiliser d'abord la méthode **BeginTrans**, une erreur se produit.</span><span class="sxs-lookup"><span data-stu-id="1954d-134">If you use the **CommitTrans** or **Rollback** method without first using the **BeginTrans** method, an error occurs.</span></span>

<span data-ttu-id="1954d-135">Certaines bases de données ISAM utilisées dans un espace de travail Microsoft Access peuvent ne pas prendre en charge les transactions, auquel cas la propriété **Transactions** de l'objet **Database** ou **Recordset** prend la valeur **False**.</span><span class="sxs-lookup"><span data-stu-id="1954d-135">Some ISAM databases used in a Microsoft Access workspace may not support transactions, in which case the **Transactions** property of the **Database** object or **Recordset** object is **False**.</span></span> <span data-ttu-id="1954d-136">Pour s'assurer que les bases de données prennent en charge les transactions, vérifiez la valeur de la propriété **Transactions** de l'objet **Database** avant d'utiliser la méthode **BeginTrans**.</span><span class="sxs-lookup"><span data-stu-id="1954d-136">To make sure the database supports transactions, check the value of the **Transactions** property of the **Database** object before using the **BeginTrans** method.</span></span> <span data-ttu-id="1954d-137">Si vous utilisez un objet **Recordset** basé sur plusieurs bases de données, vérifiez la propriété **Transactions** de l'objet **Recordset**.</span><span class="sxs-lookup"><span data-stu-id="1954d-137">If you are using a **Recordset** object based on more than one database, check the **Transactions** property of the **Recordset** object.</span></span> 

<span data-ttu-id="1954d-138">Si un objet **Recordset** est basé entièrement sur des tables de moteur de base de données Microsoft Access, vous pouvez toujours utiliser des transactions.</span><span class="sxs-lookup"><span data-stu-id="1954d-138">If a **Recordset** is based entirely on Microsoft Access database engine tables, you can always use transactions.</span></span> <span data-ttu-id="1954d-139">En revanche, les objets **Recordset** basés sur des tables créées par d'autres produits de base de données ne peuvent pas prendre en charge les transactions.</span><span class="sxs-lookup"><span data-stu-id="1954d-139">**Recordset** objects based on tables created by other database products, however, may not support transactions.</span></span> <span data-ttu-id="1954d-140">Par exemple, vous ne pouvez pas utiliser de transactions dans un objet **Recordset** basé sur une table Paradox.</span><span class="sxs-lookup"><span data-stu-id="1954d-140">For example, you can't use transactions in a **Recordset** based on a Paradox table.</span></span> <span data-ttu-id="1954d-141">Dans ce cas, la propriété **Transactions** prend la valeur **False**.</span><span class="sxs-lookup"><span data-stu-id="1954d-141">In this case, the **Transactions** property is **False**.</span></span> <span data-ttu-id="1954d-142">Si l'objet **Database** ou **Recordset** ne prend pas en charge les transactions, les méthodes sont ignorées et aucune erreur ne se produit.</span><span class="sxs-lookup"><span data-stu-id="1954d-142">If the **Database** or **Recordset** doesn't support transactions, the methods are ignored and no error occurs.</span></span>

<span data-ttu-id="1954d-143">Vous ne pouvez pas imbriquer des transactions si vous accédez à des sources de données ODBC par le biais du moteur de base de données Microsoft Access.</span><span class="sxs-lookup"><span data-stu-id="1954d-143">You can't nest transactions if you are accessing ODBC data sources through the Microsoft Access database engine.</span></span>

<span data-ttu-id="1954d-p108">Dans les espaces de travail ODBC, lorsque vous utilisez **CommitTrans**, il peut arriver que le curseur ne soit plus valide. Utilisez la méthode **Requery** pour afficher les modifications dans l'objet **Recordset**, ou fermez l'objet **Recordset** et réouvrez-le.</span><span class="sxs-lookup"><span data-stu-id="1954d-p108">In ODBC workspaces, when you use **CommitTrans** your cursor may no longer be valid. Use the **Requery** method to view the changes in the **Recordset**, or close and re-open the **Recordset**.</span></span>

> [!NOTE]
> - <span data-ttu-id="1954d-p109">Vous pouvez souvent améliorer les performances de l'application en fractionnant les opérations qui doivent accéder au disque en blocs de transactions. Cette opération place les opérations dans la mémoire tampon et peut considérablement réduire le nombre d'accès à un disque.</span><span class="sxs-lookup"><span data-stu-id="1954d-p109">You can often improve the performance of your application by breaking operations that require disk access into transaction blocks. This buffers your operations and may significantly reduce the number of times a disk is accessed.</span></span>
> - <span data-ttu-id="1954d-p110">Dans un espace de travail Microsoft Access, les transactions sont journalisées dans un fichier conservé dans le répertoire spécifié par la variable d'environnement TEMP sur le poste de travail. Si le fichier journal des transactions utilise tout l'espace de stockage disponible sur le lecteur TEMP, le moteur de base de données déclenche une erreur d'exécution. À ce niveau, si vous utilisez **CommitTrans**, un nombre indéterminé d'opérations est validé, mais les opérations restantes qui n'ont pas été terminées sont perdues, et l'opération doit être recommencée. L'utilisation de la méthode **Rollback** libère le journal des transactions et annule toutes les opérations dans la transaction.</span><span class="sxs-lookup"><span data-stu-id="1954d-p110">In a Microsoft Access workspace, transactions are logged in a file kept in the directory specified by the TEMP environment variable on the workstation. If the transaction log file exhausts the available storage on your TEMP drive, the database engine triggers a run-time error. At this point, if you use **CommitTrans**, an indeterminate number of operations are committed, but the remaining uncompleted operations are lost, and the operation has to be restarted. Using a **Rollback** method releases the transaction log and rolls back all operations in the transaction.</span></span>
> - <span data-ttu-id="1954d-152">La fermeture d'un objet **Recordset** cloné pendant une transaction en attente entraîne une opération **Rollback** implicite.</span><span class="sxs-lookup"><span data-stu-id="1954d-152">Closing a clone **Recordset** within a pending transaction will cause an implicit **Rollback** operation.</span></span>

## <a name="example"></a><span data-ttu-id="1954d-153">Exemple</span><span class="sxs-lookup"><span data-stu-id="1954d-153">Example</span></span>

<span data-ttu-id="1954d-154">L’exemple suivant montre comment utiliser une transaction dans un espace de travail Data Access Objects (DAO).</span><span class="sxs-lookup"><span data-stu-id="1954d-154">The following example shows how to use a transaction in a Data Access Objects (DAO) workspace.</span></span>

<span data-ttu-id="1954d-155">**Exemple de code fourni par** la [référence du programmeur Microsoft Access 2010](https://www.amazon.com/Microsoft-Access-2010-Programmers-Reference/dp/8126528125).</span><span class="sxs-lookup"><span data-stu-id="1954d-155">**Sample code provided by** the [Microsoft Access 2010 Programmer’s Reference](https://www.amazon.com/Microsoft-Access-2010-Programmers-Reference/dp/8126528125).</span></span>


```vb
    Public Sub TransferFunds()
        Dim wrk As DAO.Workspace
        Dim dbC As DAO.Database
        Dim dbX As DAO.Database
        
        Set wrk = DBEngine(0)
        Set dbC = CurrentDb
        Set dbX = wrk.OpenDatabase("e:\books\acc2007vba\myDB.accdb")
        
        On Error GoTo trans_Err
        
        'Begin the transaction
        
        wrk.BeginTrans
        
        'Withdraw funds from one account table
        dbC.Execute "INSERT INTO tblAccounts ( Amount, Txn, TxnDate ) SELECT -20, 'DEBIT', Date()", dbFailOnError
    
        'Deposit funds into another account table
        dbX.Execute "INSERT INTO tblAccounts ( Amount, Txn, TxnDate ) SELECT 20, 'CREDIT', Date()", dbFailOnError
        
        'Commit the transaction
        wrk.CommitTrans dbForceOSFlush
        
    trans_Exit:
        'Clean up
        wrk.Close
        Set dbC = Nothing
        Set dbX = Nothing
        Set wrk = Nothing
        Exit Sub
        
    trans_Err:
        'Roll back the transaction
        wrk.Rollback
        Resume trans_Exit
        
    End Sub
```