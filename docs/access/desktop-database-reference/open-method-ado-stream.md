---
title: Open, méthode (flux ADO)
TOCTitle: Open method (ADO Stream)
ms:assetid: fa2e6aaa-e9f5-009c-f3a0-050a00abf9b0
ms:mtpsurl: https://msdn.microsoft.com/library/JJ250275(v=office.15)
ms:contentKeyID: 48548833
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: 3a943209ce329d59fb4846ed18fd008bc45803da
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32288384"
---
# <a name="open-method-ado-stream"></a>Open, méthode (flux ADO)


**S’applique à** : Access 2013, Office 2013


Ouvre un objet [Stream](stream-object-ado.md) pour manipuler des flux de données binaires ou de texte.

## <a name="syntax"></a>Syntaxe

*Stream*. Open *Source*, *Mode*, *OpenOptions*, *UserName*, *Password*

## <a name="parameters"></a>Paramètres

|Paramètre|Description|
|:--------|:----------|
|*Source* |Facultatif. Valeur de type **Variant** qui spécifie la source de données de l'objet **Stream**. *La source* peut contenir une chaîne d’URL absolue qui pointe vers un nœud existant dans une arborescence connue, comme un système de fichiers ou de messagerie. Une URL doit être spécifiée à l’aide du mot clé d’URL ( « URL=*schéma*://*dossier* / *serveur*»). *Source* peut également contenir une référence à un objet [Record](record-object-ado.md) déjà ouvert, qui ouvre le flux par défaut associé à l’objet **Record**. Si *Source* n’est pas spécifié, un objet **Stream** est instancié et ouvert, sans association par défaut à une source sous-jacente. Pour plus d’informations sur les schémas d’URL et leurs fournisseurs associés, voir [URL absolues et relatives.](absolute-and-relative-urls.md)|
|*Mode* |Facultatif. Valeur [ConnectModeEnum](connectmodeenum.md) qui spécifie le mode d’accès de l’objet **Stream** résultant (par exemple, lecture/écriture ou en lecture seule). La valeur par défaut est **adModeUnknown**. Pour plus d’informations sur les modes d’accès, consultez la propriété [Mode](mode-property-ado.md). Si *Mode* n’est pas spécifié, il est hérité de l’objet source. Par exemple, si l’objet **Record** source est ouvert en lecture seule, l’objet **Stream** sera également ouvert en lecture seule par défaut.|
|*OpenOptions* |Facultatif. Valeur [StreamOpenOptionsEnum](streamopenoptionsenum.md). La valeur par défaut est **adOpenStreamUnspecified**.|
|*UserName* |Facultatif. Valeur de type **String** contenant l'ID utilisateur qui, le cas échéant, accède à l'objet **Stream**.|
|*MotDePasse* |Facultatif. Valeur de type **String** contenant le mot de passe qui, le cas échéant, accède à l'objet **Stream**.|

## <a name="remarks"></a>Remarques

**Lorsqu’un objet Record** est transmis en tant que paramètre source, les paramètres *UserID* et *Password* ne sont pas utilisés car l’accès à l’objet **Record** est déjà disponible. De même, le [mode de](mode-property-ado.md) l’objet **Record** est transféré vers **l’objet Stream.** Lorsque *la source* n’est pas spécifiée, le **flux** ouvert ne contient aucune donnée et a la [taille](https://docs.microsoft.com/office/vba/access/concepts/miscellaneous/size-property-ado-stream) zéro (0). To avoid losing any data that is written to this **Stream** when the **Stream** is closed, save the **Stream** with the [CopyTo](copyto-method-ado.md) or [SaveToFile](savetofile-method-ado.md) methods, or save it to another memory location.

Si *OptionsOuverture* a la valeur **adOpenStreamFromRecord**, le contenu du paramètre *Source* est identifié comme un objet **Record** déjà ouvert. Le comportement par défaut consiste à traiter *Source* comme une URL qui pointe directement vers un nœud d’une arborescence, par exemple un fichier. Le flux par défaut associé à ce nœud est ouvert.

While the **Stream** is not open, it is possible to read all the read-only properties of the **Stream**. If a **Stream** is opened asynchronously, all subsequent operations (other than checking the [State](state-property-ado.md) and other read-only properties) are blocked until the **Open** operation is completed.

Outre les options présentées ci-dessus, l’omission de *Source* vous permet d’instancier simplement un objet **Stream** en mémoire sans l’associer à une source sous-jacente. Vous pouvez ajouter dynamiquement des données à un flux en écrivant des données binaires ou de texte dans l’objet **Stream** avec les méthodes [Write](write-method-ado.md) ou [WriteText](writetext-method-ado.md), ou en chargeant les données d’un fichier à l’aide de la méthode [LoadFromFile](loadfromfile-method-ado.md).

