---
title: Charset, propriété (ADO)
TOCTitle: Charset Property (ADO)
ms:assetid: 454f664e-6d62-eec9-487d-882c2f9503b0
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249213(v=office.15)
ms:contentKeyID: 48544551
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 84436f814bbc2a503e843d4e9832fa2bab612b58
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471832"
---
# <a name="charset-property-ado"></a><span data-ttu-id="6e3c8-102">Charset, propriété (ADO)</span><span class="sxs-lookup"><span data-stu-id="6e3c8-102">Charset Property (ADO)</span></span>


<span data-ttu-id="6e3c8-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="6e3c8-103">**Applies to**: Access 2013 | Office 2013</span></span>

<span data-ttu-id="6e3c8-104">Indique le jeu de caractères dans lequel le contenu d'un objet [Stream](stream-object-ado.md) de type texte doit être converti pour être stocké dans la mémoire tampon interne des objets Stream.</span><span class="sxs-lookup"><span data-stu-id="6e3c8-104">Indicates the character set into which the contents of a text [Stream](stream-object-ado.md) should be translated for storage in the Stream objects internal buffer.</span></span>

## <a name="settings-and-return-values"></a><span data-ttu-id="6e3c8-105">Paramètres et valeurs de retour</span><span class="sxs-lookup"><span data-stu-id="6e3c8-105">Settings and Return Values</span></span>

<span data-ttu-id="6e3c8-106">Définit ou renvoie une valeur de type **String** qui spécifie le jeu de caractères dans lequel le contenu de l'objet **Stream** est converti.</span><span class="sxs-lookup"><span data-stu-id="6e3c8-106">Sets or returns a **String** value that specifies the character set into which the contents of the **Stream** will be translated.</span></span> <span data-ttu-id="6e3c8-107">La valeur par défaut est « Unicode ».</span><span class="sxs-lookup"><span data-stu-id="6e3c8-107">The default value is "Unicode".</span></span> <span data-ttu-id="6e3c8-108">Les valeurs admises sont des chaînes classiques transmises via l'interface sous forme de chaînes de jeu de caractères Internet (par exemple, « ISO-8859-1 », « Windows-1252 », etc.).</span><span class="sxs-lookup"><span data-stu-id="6e3c8-108">Allowed values are typical strings passed over the interface as Internet character set strings (for example, "iso-8859-1", "Windows-1252", etc.).</span></span> <span data-ttu-id="6e3c8-109">Pour obtenir la liste des chaînes de jeu de caractères qui est appelée par un système, voir les sous-clés de HKEY\_CLASSES\_racine\\MIME\\base de données\\jeu de caractères dans le Registre Windows.</span><span class="sxs-lookup"><span data-stu-id="6e3c8-109">For a list of the character set strings that is known by a system, see the subkeys of HKEY\_CLASSES\_ROOT\\MIME\\Database\\Charset in the Windows Registry.</span></span>

## <a name="remarks"></a><span data-ttu-id="6e3c8-110">Notes</span><span class="sxs-lookup"><span data-stu-id="6e3c8-110">Remarks</span></span>

<span data-ttu-id="6e3c8-p102">Dans un objet **Stream** de type texte, les données de texte sont stockées dans le jeu de caractères spécifié par la propriété **Charset**. La valeur par défaut est Unicode. La propriété **Charset** est utilisée pour convertir des données accédant à l'objet **Stream** ou provenant de l'objet **Stream**. Par exemple, si l'objet **Stream** contient des données au format ISO-8859-1 et que ces données sont copiées dans un BSTR, l'objet **Stream** convertit les données au format Unicode. L'inverse est vrai également.</span><span class="sxs-lookup"><span data-stu-id="6e3c8-p102">In a text **Stream** object, text data is stored in the character set specified by the **Charset** property. The default is Unicode. The **Charset** property is used for converting data going into the **Stream** or coming out of the **Stream**. For example, if the **Stream** contains ISO-8859-1 data and that data is copied to a BSTR, the **Stream** object will convert the data to Unicode. The reverse is also true.</span></span>

<span data-ttu-id="6e3c8-116">Pour un objet **Stream** ouvert, la propriété [Position](position-property-ado.md) actuelle doit se trouver au début de l'objet **Stream** (0) pour pouvoir définir la propriété **Charset**.</span><span class="sxs-lookup"><span data-stu-id="6e3c8-116">For an open **Stream**, the current [Position](position-property-ado.md) must be at the beginning of the **Stream** (0) to be able to set **Charset**.</span></span>

<span data-ttu-id="6e3c8-p103">La propriété **Charset** n’est utilisée qu’avec des objets **Stream** de type texte ([Type](type-property-ado-stream.md) a la valeur **adTypeText**). Cette propriété est ignorée si **Type** a la valeur **adTypeBinary**.</span><span class="sxs-lookup"><span data-stu-id="6e3c8-p103">**Charset** is used only with text **Stream** objects ([Type](type-property-ado-stream.md) is **adTypeText**). This property is ignored if **Type** is **adTypeBinary**.</span></span>
