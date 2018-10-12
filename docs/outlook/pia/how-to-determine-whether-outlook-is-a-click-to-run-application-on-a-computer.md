---
title: Vérifier si Outlook est une application « Démarrer en un clic » sur un ordinateur
TOCTitle: Determine whether Outlook is a Click-to-Run application on a computer
ms:assetid: 1b8573be-8ea8-4973-869d-87fda57ce525
ms:mtpsurl: https://msdn.microsoft.com/library/Ff522355(v=office.15)
ms:contentKeyID: 55119804
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: c155fed72a314c5c260ed2fe574474afd45c44c3
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25406651"
---
# <a name="determine-whether-outlook-is-a-click-to-run-application-on-a-computer"></a><span data-ttu-id="0e258-102">Vérifier si Outlook est une application « Démarrer en un clic » sur un ordinateur</span><span class="sxs-lookup"><span data-stu-id="0e258-102">Determine whether Outlook is a Click-to-Run application on a computer</span></span>

<span data-ttu-id="0e258-103">« Démarrer en un clic » est un mécanisme de livraison et de mise à jour de logiciels disponible sur Office 2010 et versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="0e258-103">Click-to-Run is a software delivery and updating mechanism available to Microsoft Office 2010.</span></span> <span data-ttu-id="0e258-104">Les produits livrés au moyen de « Démarrer en un clic » s’exécutent dans un environnement d’applications virtuelles sur le système d’exploitation local.</span><span class="sxs-lookup"><span data-stu-id="0e258-104">Products delivered via Click-to-Run execute in a virtual application environment on the local operating system.</span></span> <span data-ttu-id="0e258-105">Cela signifie qu’ils ont des copies privées de leurs fichiers et paramètres, et que les modifications qu’ils apportent sont capturées dans l’environnement virtuel.</span><span class="sxs-lookup"><span data-stu-id="0e258-105">This means that they have private copies of their files and settings, and that any changes they make are captured in the virtual environment.</span></span>

<span data-ttu-id="0e258-106">Click-to-Run est rapide, les utilisateurs peuvent commencer rapidement à se servir d'une application sans devoir attendre l'installation du produit complet.</span><span class="sxs-lookup"><span data-stu-id="0e258-106">Click-to-Run is fast—users can start running an application within a short time without waiting for the complete product to finish installing.</span></span> <span data-ttu-id="0e258-107">Les mises à jour s'exécutent automatiquement en arrière-plan, sans que l'utilisateur ne doive d'abord retirer une installation ni procéder manuellement à l'installation des mises à jour.</span><span class="sxs-lookup"><span data-stu-id="0e258-107">Updates are run automatically in the background, without requiring a user to first remove an installation or manually install updates.</span></span> <span data-ttu-id="0e258-108">Les produits « Démarrer en un clic » sont virtualisés et ne sont pas en conflit avec d’autres logiciels installés.</span><span class="sxs-lookup"><span data-stu-id="0e258-108">Click-to-Run products are virtualized and do not conflict with other installed software.</span></span> <span data-ttu-id="0e258-109">Cependant, comme un produit livré par « Démarrer en un clic » a des copies privées de tous ses fichiers et de son inscription, un développeur de complément ne peut pas déterminer l’existence du produit de la même manière que pour un produit installé sur le disque dur d’un ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="0e258-109">However, because a product delivered by Click-to-Run has private copies of all its files and registration, an add-in developer cannot determine the product's existence in the same manner as a product that was installed on a client computer's hard disk.</span></span> <span data-ttu-id="0e258-110">À partir d’Outlook 2010, les développeurs de compléments doivent déterminer si Outlook a été installé, et si Outlook a été fourni sous la forme d’un produit « Démarrer en un clic ».</span><span class="sxs-lookup"><span data-stu-id="0e258-110">Starting with Outlook 2010, add-in developers should verify whether Outlook has been installed, and whether Outlook has been delivered as a Click-to-Run product.</span></span>


> [!NOTE]
> <span data-ttu-id="0e258-111">Seul Office 2013 version 32 bits est pris en charge dans l’environnement d’applications virtuelles « Démarrer en un clic », même si l’ordinateur client utilise un système d’exploitation 64 bits.</span><span class="sxs-lookup"><span data-stu-id="0e258-111">Only 32-bit Office 2010 is supported in the Click-to-Run virtual application environment, even if the client computer runs a 64-bit operating system.</span></span>



### <a name="to-check-whether-outlook-2013-was-delivered-by-click-to-run-on-a-client-computer"></a><span data-ttu-id="0e258-112">Pour vérifier si Outlook 2013 a été fourni par « Démarrer en un clic » sur un ordinateur client, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="0e258-112">To check whether Outlook was delivered by Click-to-Run on a client computer</span></span>

- <span data-ttu-id="0e258-113">Vérifiez si la clé VirtualOutlook existe dans l’emplacement suivant du Registre Windows :</span><span class="sxs-lookup"><span data-stu-id="0e258-113">Verify whether the VirtualOutlook key exists in the following location in the Windows registry:</span></span>
    
  `HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Common\InstallRoot\Virtual\VirtualOutlook`
    
  <span data-ttu-id="0e258-114">La clé VirtualOutlook est une valeur REG\_SZ qui contient la balise culture de la langue installée du produit, par exemple « en-us ».</span><span class="sxs-lookup"><span data-stu-id="0e258-114">
The VirtualOutlook key is a REG_SZ value that contains the culture tag of the installed product language, such as "en-us". 
</span></span>

## <a name="see-also"></a><span data-ttu-id="0e258-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0e258-115">See also</span></span>

- [<span data-ttu-id="0e258-116">Administration des compléments</span><span class="sxs-lookup"><span data-stu-id="0e258-116">Add-in Administration</span></span>](add-in-administration.md)
