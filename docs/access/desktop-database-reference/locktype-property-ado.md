---
title: LockType, propriété (ADO)
TOCTitle: LockType Property (ADO)
ms:assetid: 1d2622dc-6cab-1b7f-98a8-97a41d5c047f
ms:mtpsurl: https://msdn.microsoft.com/library/JJ248965(v=office.15)
ms:contentKeyID: 48543589
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 6d1740f42fae3485d88a4820081621f7f2483c51
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471726"
---
# <a name="locktype-property-ado"></a><span data-ttu-id="d5c3d-102">LockType, propriété (ADO)</span><span class="sxs-lookup"><span data-stu-id="d5c3d-102">LockType Property (ADO)</span></span>


<span data-ttu-id="d5c3d-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="d5c3d-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="d5c3d-104">Indique le type des verrous placés sur les enregistrements pendant leur modification.</span><span class="sxs-lookup"><span data-stu-id="d5c3d-104">Indicates the type of locks placed on records during editing.</span></span>

## <a name="settings-and-return-values"></a><span data-ttu-id="d5c3d-105">Paramètres et valeurs renvoyées</span><span class="sxs-lookup"><span data-stu-id="d5c3d-105">Settings and Return Values</span></span>

<span data-ttu-id="d5c3d-p101">Définit ou renvoie une valeur [LockTypeEnum](locktypeenum.md). La valeur par défaut est **adLockReadOnly**.</span><span class="sxs-lookup"><span data-stu-id="d5c3d-p101">Sets or returns a [LockTypeEnum](locktypeenum.md) value. The default value is **adLockReadOnly**.</span></span>

## <a name="remarks"></a><span data-ttu-id="d5c3d-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="d5c3d-108">Remarks</span></span>

<span data-ttu-id="d5c3d-p102">Définissez la propriété **LockType** avant d'ouvrir un [Recordset](recordset-object-ado.md) pour spécifier le type de verrouillage que le fournisseur doit utiliser pour l'ouvrir. Lisez la propriété pour connaître le type de verrouillage utilisé sur un objet **Recordset** ouvert.</span><span class="sxs-lookup"><span data-stu-id="d5c3d-p102">Set the **LockType** property before opening a [Recordset](recordset-object-ado.md) to specify what type of locking the provider should use when opening it. Read the property to return the type of locking in use on an open **Recordset** object.</span></span>

<span data-ttu-id="d5c3d-p103">Les fournisseurs peuvent ne pas prendre en charge tous types de verrouillage. Si un fournisseur ne reconnaît pas le paramètre **LockType** demandé, il le remplacera par un autre type de verrouillage. Pour déterminer la fonctionnalité réelle de verrouillage disponible pour un objet **Recordset**, utilisez la méthode [Supports](supports-method-ado.md) avec **adUpdate** et **adUpdateBatch**.</span><span class="sxs-lookup"><span data-stu-id="d5c3d-p103">Providers may not support all lock types. If a provider cannot support the requested **LockType** setting, it will substitute another type of locking. To determine the actual locking functionality available in a **Recordset** object, use the [Supports](supports-method-ado.md) method with **adUpdate** and **adUpdateBatch**.</span></span>

<span data-ttu-id="d5c3d-p104">Le paramètre **adLockPessimistic** n'est pas pris en charge si la valeur de la propriété [CursorLocation](cursorlocation-property-ado.md) est **adUseClient**. Si la valeur spécifiée n'est pas prise en charge, aucune erreur n'est générée ; c'est le **LockType** pris en charge le plus proche qui est utilisé.</span><span class="sxs-lookup"><span data-stu-id="d5c3d-p104">The **adLockPessimistic** setting is not supported if the [CursorLocation](cursorlocation-property-ado.md) property is set to **adUseClient**. If an unsupported value is set, then no error will result; the closest supported **LockType** will be used instead.</span></span>

<span data-ttu-id="d5c3d-116">La propriété **LockType** est accessible en lecture et en écriture lorsque le **Recordset** est fermé et en lecture seule lorsqu' il est ouvert.</span><span class="sxs-lookup"><span data-stu-id="d5c3d-116">The **LockType** property is read/write when the **Recordset** is closed and read-only when it is open.</span></span>

<span data-ttu-id="d5c3d-117">**L’utilisation du Service de données à distance** Lorsqu’elle est utilisée sur un objet Recordset de côté client, la propriété **LockType** peut uniquement être définie **: adLockBatchOptimistic**.</span><span class="sxs-lookup"><span data-stu-id="d5c3d-117">**Remote Data Service Usage**When used on a client-side Recordset object, the **LockType** property can only be set to **adLockBatchOptimistic**.</span></span>
