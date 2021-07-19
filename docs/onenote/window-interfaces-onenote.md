---
title: Interfaces de fenêtre (OneNote 2013)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: overview
localization_priority: Normal
ms.assetid: e6900ad7-c147-4816-93a9-5773170b115a
description: Les interfaces Window et Windows sont que des objets OneNote 2013 API qui permet aux utilisateurs de travailler avec les fenêtres OneNote. Ces objets permettent aux utilisateurs d'énumérer l'ensemble des fenêtres OneNote et de modifier certaines propriétés de la fenêtre.
ms.openlocfilehash: efc34312def588ecff54c63b3db84f8bf909352b
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32317033"
---
# <a name="window-interfaces-onenote"></a>Interfaces de fenêtre (OneNote 2013)

Les interfaces **Window** et **Windows** sont des objets API OneNote 2013 qui permettent aux utilisateurs de travailler avec les fenêtres OneNote. Ces objets permettent aux utilisateurs d'énumérer l'ensemble des fenêtres OneNote et de modifier certaines propriétés de la fenêtre. 
  
## <a name="onenote-window-views"></a>Affichages de fenêtre OneNote

La liste suivante indique les modes de quatre affichage que vous pouvez utiliser pour les fenêtres OneNote : 
  
- Mode normal—affiche la fenêtre de OneNote par défaut dans lequel les volets de navigation Bloc-notes et de Page sont visibles.
    
- Page vue d'ensemble—affiche une minimale-interface utilisateur (IU) de vue dans laquelle les volets bloc-notes et de Page ne sont pas affichées.
    
- Note rapide—affiche une fenêtre de petite taille qui permet aux utilisateurs de prendre des notes courtes. Vous accédez généralement notes rapides par le biais de l'icône de OneNote dans la zone de notification de Windows, mais vous pouvez également accéder par le biais de l'onglet **affichage** dans OneNote. 
    
- Ancrer sur le bureau—affiche une fenêtre OneNote que vous pouvez ancrer à n'importe quel côté du bureau (similaire à la barre des tâches). Cet affichage réduit la taille du bureau pour s'adapter à la fenêtre. Vous pouvez ancrer qu'une seule fenêtre à tout moment, et la fenêtre est toujours visible sans le blocage du bureau. 
    
La figure suivante montre l’apparence de l’affichage Page complète, De la station d’accueil au bureau et des notes rapides sur votre bureau.
  
**Affichages OneNote**

![OneNote fenêtre et les](media/ON15Con_views.jpg "OneNote fenêtre")
  
## <a name="interfaces"></a>Interfaces

Cette section répertorie les interfaces et les membres que vous pouvez utiliser pour modifier par programme les fenêtres OneNote.
  
### <a name="windows-interface"></a>Windows interface

L'interface **Windows** permet à l'utilisateur d'accéder à l'ensemble des fenêtres ouvertes de OneNote. Il s'agit d'une propriété de la classe **Application** OneNote, accédée via **Application.Windows**. Ce qui retourne l'ensemble énuméré des fenêtres OneNote. 
  
**Propriétés**

|**Name**|**Type**|**Description**|
|:-----|:-----|:-----|
|**Count** <br/> |ulong  <br/> |Obtient le nombre d'objets de **Window** dans le jeu de **Windows**.  <br/> |
|**CurrentWindow** <br/> |**Window** <br/> |Obtient l'objet **Window** de la fenêtre OneNote active.  <br/> |
|**Items** <br/> |**Window** <br/> |Renvoie l'objet **Window** qui correspond à la valeur d'index transmise. Cette propriété n'est pas accessible directement. Pour renvoyer un objet **Window**, utilisez **Windows [(uint) index]**. <br/> |
   
### <a name="window-interface"></a>Interface de fenêtre

L'interface **Window** permet à l'utilisateur d'accéder à certaines propriétés de chaque fenêtre. Chaque fenêtre OneNote est accessible par l'énumération par le biais de la propriété **Windows** de la classe **Application**. 
  
**Propriétés**

|**Name**|**Type**|**Description**|
|:-----|:-----|:-----|
|**Active** <br/> |bool  <br/> |Obtient ou définit une valeur qui indique si la fenêtre est la fenêtre active de OneNote.  <br/> |
|**Application** <br/> |**Application** <br/> |Obtient l'objet **Application** OneNote qui est associé à la fenêtre.  <br/> |
|**CurrentPageId** <br/> |string  <br/> |Obtient l'ID d'objet de la page OneNote active de la fenêtre.  <br/> |
|**CurrentSectionId** <br/> |string  <br/> |Obtient l'ID d'objet de la section OneNote active de la fenêtre.  <br/> |
|**CurrentSectionGroupId** <br/> |string  <br/> |Obtient l'ID d'objet du groupe de section OneNote active de la fenêtre.  <br/> |
|**CurrentNotebookId** <br/> |string  <br/> |Obtient l'ID d'objet du bloc-notes OneNote actif de la fenêtre.  <br/> |
|**DockedLocation** <br/> |**DockedLocation** <br/> |Obtient ou définit l'emplacement d'ancrage de la fenêtre OneNote.  <br/> |
|**FullPageView** <br/> |bool  <br/> |Obtient ou définit une valeur qui indique si la fenêtre est en mode Page entière (affichage de l'interface utilisateur minimale).  <br/> |
|**SideNote** <br/> |bool  <br/> |Obtient ou définit une valeur qui indique si la fenêtre est une fenêtre de note rapide.  <br/> |
|**WindowHandle** <br/> |ulong  <br/> |Obtient l'ID de handle de la fenêtre OneNote.  <br/> |
   
**Méthodes**
  
Vous pouvez utiliser les méthodes suivantes de l'interface **Window** pour naviguer vers les objets spécifiés dans la fenêtre OneNote ou à l'URL spécifiées. 
  
**NavigateTo**

|||
|:-----|:-----|
|**Description** <br/> |Accède à l'objet spécifié dans la fenêtre OneNote. Par exemple, vous pouvez naviguer vers des sections, des pages et des éléments du plan au sein de pages.  <br/> |
|**Syntaxe** <br/> | `HRESULT NavigateTo(`           ` [in]BSTR bstrHierarchyObjectID, `           ` [in]BSTR bstrObjectID); ` <br/> |
|**Paramètres** <br/> | _bstrHierarchyObjectID_—la hiérarchie OneNote des ID de l'objet que vous voulez atteindre. L'ID d'objet peut faire référence à un bloc-notes OneNote, section, le groupe de section ou page. <br/>  _bstrObjectID_—l'ID de OneNote de l'objet spécifique à accéder à l'intérieur d'une page OneNote. Si l'utilisateur ne souhaite pas d'accéder à un objet spécifique dans une page, ce paramètre est défini sur null. <br/> |
   
**NavigateToUrl**

|||
|:-----|:-----|
|**Description** <br/> |Si un lien OneNote est transmis (onenote://), cette méthode ouvre la fenêtre OneNote à l’emplacement correspondant dans OneNote. Toutefois, si le lien est un lien externe, tel https:// ou file://, une boîte de dialogue de sécurité s’affiche. Après le renvoi, OneNote essaie d'ouvrir la liaison et une erreur HResult.hrObjectDoesNotExist est renvoyée.  <br/> |
|**Syntaxe** <br/> | `HRESULT NavigateToUrl (`           ` [in]BSTR bstrUrl); ` <br/> |
|**Paramètres** <br/> | _bstrUrl_— l'URL à atteindre.  <br/> |
   
**SetDockedLocation**

|||
|:-----|:-----|
|**Description** <br/> |Ancre la fenêtre à l'emplacement spécifié par **dockLocation** et le moniteur au **ptMonitor**.  <br/> |
|**Syntaxe** <br/> | `HRESULT SetDockedLocation`(           `[in] DockLocation dockLocation,`           `[in] POINT ptMonitor);` <br/> |
|**Paramètres** <br/> | _dockLocation_ - indique l'emplacement d'ancrage d'une fenêtre OneNote 2013.  <br/>  _ptMonitor_ - indique (facultatif) dans x, y coordonnées dont la fenêtre de surveillance doit être ancré dans.  <br/> |
   
## <a name="example"></a>Exemple

Le code suivant effectue une itération dans les fenêtres OneNote pour rechercher une fenêtre fixe. Si aucune fenêtre ancrée n'existe, l'exemple ancre la fenêtre active. Si aucune fenêtre active n'existe, le code crée une nouvelle fenêtre ancrée.
  
```cs
using System;
using System.Diagnostics;
using Microsoft.Office.Interop.OneNote;
namespace SampleWND
{
    class DockOneNoteWindow
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.OneNote.Application app = new Microsoft.Office.Interop.OneNote.Application();
            // Search through all OneNote windows for a docked window and activate it.
            bool foundDockedWND = false;
            for (int i = 0; i < app.Windows.Count; i++)
            {
                if (app.Windows[(uint) i].DockedLocation != DockLocation.dlNone)
                {
                    foundDockedWND = true;
                    app.Windows[(uint) i].Active = true;
                }
            }
            
            // If no docked window exists, dock the active window.
            if (!foundDockedWND && (app.Windows.Count > 0))
                app.Windows.CurrentWindow.DockedLocation = DockLocation.dlDefault;
            // If no active window exists, create a new docked window.
            if (app.Windows.Count < 1)
            {
                Process oneProc = new Process();
                oneProc.StartInfo.FileName = "onenote.exe";
                oneProc.StartInfo.Arguments = "/docked";
                oneProc.Start();
            }
        }
    }
}

```

## <a name="see-also"></a>Voir aussi

- [Référence du développeur OneNote](onenote-developer-reference.md)

