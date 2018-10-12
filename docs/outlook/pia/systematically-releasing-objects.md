---
title: Lancement systématique des objets
TOCTitle: Systematically releasing objects
ms:assetid: d4cd1d8e-aae6-483b-a4d8-1656171e838d
ms:mtpsurl: https://msdn.microsoft.com/library/Bb623945(v=office.15)
ms:contentKeyID: 55119785
ms.date: 07/24/2014
mtps_version: v=office.15
ms.openlocfilehash: 812f720b3ebab1100040802d7593548063497297
ms.sourcegitcommit: ef717c65d8dd41ababffb01eafc443c79950aed4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/04/2018
ms.locfileid: "25406833"
---
# <a name="systematically-releasing-objects"></a><span data-ttu-id="f006a-102">Lancement systématique des objets</span><span class="sxs-lookup"><span data-stu-id="f006a-102">Systematically Releasing Objects</span></span>

<span data-ttu-id="f006a-103">Cette rubrique fait la synthèse des recommandations sur l'arrêt des compléments pour les développeurs de compléments et les administrateurs qui utilisent Outlook.</span><span class="sxs-lookup"><span data-stu-id="f006a-103">This topic summarizes add-in shutdown recommendations for add-in developers and IT administrators that use Microsoft Outlook 2010.</span></span> <span data-ttu-id="f006a-104">Pour plus d’informations, voir[arrêt des modifications pour Outlook 2010](https://msdn.microsoft.com/library/ee720183\(v=office.15\)).</span><span class="sxs-lookup"><span data-stu-id="f006a-104">For more information, see [Shutdown Changes for Outlook 2010](https://msdn.microsoft.com/library/ee720183\(v=office.15\)).</span></span>

## <a name="add-in-shutdown-changes-in-outlook"></a><span data-ttu-id="f006a-105">L’application shutdown change dans Outlook</span><span class="sxs-lookup"><span data-stu-id="f006a-105">Add-in shutdown changes in Outlook</span></span>

<span data-ttu-id="f006a-106">En commençant dans Outlook 2010, Outlook, par défaut, ne signale pas de compléments qu’il arrête.</span><span class="sxs-lookup"><span data-stu-id="f006a-106">Starting in Outlook 2010, Outlook, by default, does not signal add-ins that it is shutting down.</span></span> <span data-ttu-id="f006a-107">Plus précisément, Outlook n'appelle plus les méthodes\*\* OnBeginShutdown(Matrice)\*\* et**OnDisconnectionn(ext\_DisconnectMode, Matrice)** méthodes de l'interface**IDTExtensibility2**lors des arrêts rapides.</span><span class="sxs-lookup"><span data-stu-id="f006a-107">Specifically, Outlook no longer calls the OnBeginShutdown and OnDisconnection methods of the IDTExtensibility2 interface during fast shutdown.</span></span> <span data-ttu-id="f006a-108">De façon similaire, un complément Outlook écrit avec des outils de développement Microsoft Office dans Microsoft Visual Studio 2010 ou une version plus récente, n'appelle plus la méthode ThisAddin\_Shutdown lorsqu’Outlook s'arrête.</span><span class="sxs-lookup"><span data-stu-id="f006a-108">Similarly, an Outlook add-in written with Microsoft Office development tools in Microsoft Visual Studio 2010 no longer calls the \_ method when Outlook is shutting down.</span></span> 

<span data-ttu-id="f006a-109">La raison pour laquelle ces méthodes ne sont plus appelées est que, alors qu'une majorité de compléments effectuent des tâches simples telles que libérer des références, certains compléments effectuent des appels de services Web ou d'autres opérations de longue durée en même temps que ces événements, ce qui retarde l'arrêt d'Outlook de façon significative.</span><span class="sxs-lookup"><span data-stu-id="f006a-109">The reason to stop calling these methods is that while a majority of add-ins perform simple tasks like releasing references, some add-ins make Web service calls or other long-running operations synchronously during these events, significantly delaying Outlook from shutting down.</span></span> <span data-ttu-id="f006a-110">Suite à cette modification, Outlook fonctionne mieux que dans le passé lorsque vous l’arrêtez.</span><span class="sxs-lookup"><span data-stu-id="f006a-110">As a result of this change, Outlook works better than it has in the past when shutting down.</span></span>

## <a name="recommendations-for-add-in-shutdown-for-developers"></a><span data-ttu-id="f006a-111">Recommandations pour les développeurs relatives à l'arrêt des compléments</span><span class="sxs-lookup"><span data-stu-id="f006a-111">Recommendations for Add-in Shutdown for Developers</span></span>

<span data-ttu-id="f006a-112">Il est important que les développeurs respectent les recommandations suivantes pour l'arrêt des compléments de façon à garantir que les utilisateurs finals continuent de bénéficier d'un arrêt d'Outlook rapide et réactif.</span><span class="sxs-lookup"><span data-stu-id="f006a-112">It is important that developers observe the following recommendations for add-in shutdown to ensure that end users continue to benefit from a fast and responsive Outlook shutdown experience.</span></span>

- <span data-ttu-id="f006a-113">Les compléments doivent continuer d'implémenter les méthodes OnBeginShutdown et OnDisconnection ou ThisAddin\_Shutdown pour libérer des références et de la mémoire allouée, car il existe des scénarios dans lesquels les administrateurs souhaitent revenir à un arrêt lent via la stratégie de groupe ou dans lesquels l'utilisateur peut déconnecter manuellement votre complément via la boîte de dialogue **Compléments COM**.</span><span class="sxs-lookup"><span data-stu-id="f006a-113">Add-ins should continue to implement the OnBeginShutdown and OnDisconnection methods or ThisAddin_Shutdown to release references and allocated memory, because there are scenarios in which administrators might revert to slow shutdown through group policy, or the user might manually disconnect your add-in through the COM Add-ins dialog box.</span></span>

- <span data-ttu-id="f006a-114">Les développeurs de complément doivent seulement effectuer des tâches qui doivent absolument avoir lieu pendant l’arrêt.</span><span class="sxs-lookup"><span data-stu-id="f006a-114">Add-in developers should only perform tasks that absolutely must take place during shutdown.</span></span>

- <span data-ttu-id="f006a-115">Les développeurs de compléments doivent évaluer les performances de leurs compléments dans différents scénarios et selon différents paramètres du Registre Windows contrôlant l'arrêt d'Outlook, et modifier leurs compléments de façon pro-active pour qu'ils fonctionnent avec Outlook.</span><span class="sxs-lookup"><span data-stu-id="f006a-115">Add-in developers should evaluate the performance of their add-ins in various scenarios and under different Windows registry settings that control Outlook shutdown, and proactively modify their add-ins to work with Outlook.</span></span>

## <a name="recommendations-for-add-in-shutdown-for-it-administrators"></a><span data-ttu-id="f006a-116">Recommandations pour les administrateurs relatives à l'arrêt des compléments</span><span class="sxs-lookup"><span data-stu-id="f006a-116">Recommendations for Add-in Shutdown for IT Administrators</span></span>

<span data-ttu-id="f006a-117">Pour les administrateurs, s'il existe des compléments déjà déployés dans une entreprise et qu'ils ne peuvent pas être mis à niveau pour devenir compatibles avec le nouveau mécanisme d'arrêt d'Outlook, ils peuvent recourir à quelques nouveaux paramètres du Registre Windows pour revenir au comportement d'arrêt lent.</span><span class="sxs-lookup"><span data-stu-id="f006a-117">For IT administrators, if there are add-ins that are already deployed in an enterprise and cannot be upgraded to become compatible with the new shutdown Outlook mechanism, IT administrators can resort to a couple of new Windows registry settings to revert to the slow shutdown behavior.</span></span>

### <a name="individual-add-in-setting"></a><span data-ttu-id="f006a-118">Paramétrage d'un complément individuel</span><span class="sxs-lookup"><span data-stu-id="f006a-118">Individual Add-in Setting</span></span>

<span data-ttu-id="f006a-p104">Les administrateurs peuvent permettre des notifications d'arrêt à des compléments Outlook dans le cadre d'un déploiement de compléments. Bien que vous ne puissiez pas faire cela via une stratégie de groupe, c'est pratique si vous avez des conditions requises en matière de compatibilité descendante pour des compléments spécifiques.</span><span class="sxs-lookup"><span data-stu-id="f006a-p104">IT administrators can enable shutdown notifications to individual Outlook add-ins as part of the add-in deployment. Although you cannot do this through group policy, it is useful if you have backward-compatibility requirements for specific add-ins.</span></span>

<span data-ttu-id="f006a-p105">Configurez ce paramètre pour chaque complément via l'inscription du complément dans les ruches du Registre HKCU ou HKLM, en ajoutant une valeur supplémentaire à l'inscription du complément. Entrez le texte suivant sur une seule ligne :</span><span class="sxs-lookup"><span data-stu-id="f006a-p105">Configure this setting for each add-in through the add-in registration in the HKCU or the HKLM registry hives, by adding an additional value to the add-in registration. Type the following text as a single line:</span></span>

`HKCU\Software\Microsoft\Office\Outlook\Add-ins\<ProgID>[RequireShutdownNotification]=dword:0x1`

<span data-ttu-id="f006a-123">Paramétrer cette valeur 0x1permet au complément de recevoir des rappels bloqués pendant la fermeture d’Outlook.</span><span class="sxs-lookup"><span data-stu-id="f006a-123">Setting this value to 0x1 enables the add-in to receive blocked callbacks during Outlook shutdown.</span></span> <span data-ttu-id="f006a-124">Ceci a un impact sur les performances de l'arrêt d'Outlook et doit être évalué dans le cadre d'un déploiement.</span><span class="sxs-lookup"><span data-stu-id="f006a-124">This has an impact on the performance of Outlook shutdown and should be evaluated as part of a deployment.</span></span> <span data-ttu-id="f006a-125">Ce paramètre doit être utilisé uniquement si un complément possède de problèmes de compatibilité significatifs avec le nouveau mécanisme d’arrêt.</span><span class="sxs-lookup"><span data-stu-id="f006a-125">This setting should be used only if an add-in has significant compatibility issues with the new shutdown mechanism.</span></span> <span data-ttu-id="f006a-126">Le paramétrage de la valeur à 0x0 utilise le comportement par défaut d'Outlook.</span><span class="sxs-lookup"><span data-stu-id="f006a-126">Setting the value to 0x0 uses the default behavior of Outlook 2010.</span></span>

### <a name="global-setting"></a><span data-ttu-id="f006a-127">Paramètre global</span><span class="sxs-lookup"><span data-stu-id="f006a-127">Global setting</span></span>

<span data-ttu-id="f006a-128">Les administrateurs informatiques peuvent globalement activer les notifications d’arrêt pour tous les compléments via une stratégie de groupe.</span><span class="sxs-lookup"><span data-stu-id="f006a-128">IT administrators can enable shutdown notifications globally for all add-ins through group policy.</span></span> <span data-ttu-id="f006a-129">Ceci est recommandé pour les organisations qui ont un grand nombre de solutions internes ou d'ordinateurs où ce paramétrage doit être déployé pour assurer la compatibilité lors d'une mise en œuvre d'Outlook.</span><span class="sxs-lookup"><span data-stu-id="f006a-129">This is recommended for organizations that have a large number of internal solutions or desktops that need to deploy this setting to ensure compatibility during a rollout of Outlook 2010.</span></span>

<span data-ttu-id="f006a-130">Utilisez ce paramétrage pour modifier le nouveau mécanisme d'arrêt de façon à ce qu'il corresponde à celui utilisé par Outlook 2007.</span><span class="sxs-lookup"><span data-stu-id="f006a-130">Use this setting to change the new shutdown mechanism to match that used by Outlook 2007.</span></span> <span data-ttu-id="f006a-131">Vous pouvez déployer le paramétrage de stratégie de groupe, par utilisateur ou par ordinateur.</span><span class="sxs-lookup"><span data-stu-id="f006a-131">You can deploy the setting through group policy, either per user or per computer.</span></span> <span data-ttu-id="f006a-132">Tapez le texte suivant en une seule ligne :</span><span class="sxs-lookup"><span data-stu-id="f006a-132">Type the following text as a single line:</span></span>

`HKCU\Policies\Microsoft\Office\Outlook\15.0\Options\Shutdown[AddinFastShutdownBehavior]=dword:0x1`

<span data-ttu-id="f006a-133">Le paramétrage de AddinFastShutdownBehavior à 0x1 permet les notifications d'arrêt pour tous les compléments. Le paramétrage de la valeur à 0x0 utilise le comportement par défaut d'Outlook.</span><span class="sxs-lookup"><span data-stu-id="f006a-133">Setting   to 0x1 enables shutdown notifications for all add-ins. Setting the value to 0x0 uses the default behavior of Outlook 2010.</span></span>
