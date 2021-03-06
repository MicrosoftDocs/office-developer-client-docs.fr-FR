---
title: Excel4/Excel12
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- Excel12
- Excel4
keywords:
- fonction excel4 [excel 2007],fonction Excel12 [Excel 2007]
localization_priority: Normal
ms.assetid: 2404f10d-8641-4ee6-a909-1c5a26610f80
description: 'S’applique à : Excel 2013 | Office 2013 | Visual Studio'
ms.openlocfilehash: 7c3af5f380ae4144890b1f7b486a61a05c19de74
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33429445"
---
# <a name="excel4excel12"></a>Excel4/Excel12

**S’applique à** : Excel 2013 | Office 2013 | Visual Studio 
  
Appelle une fonction de feuille Microsoft Excel interne, une fonction ou une commande de feuille macro, ou une commande ou une fonction spéciale XLL uniquement, à partir d’une DLL/XLL ou d’une ressource de code.
  
Toutes les versions récentes de Excel prise **en charge d’Excel4.** À compter Excel 2007, **Excel12** est pris en charge. 
  
Ces fonctions ne peuvent être appelées que si Excel a passé le contrôle à la DLL ou au XLL. Ils peuvent également être appelés lorsque Excel a passé le contrôle indirectement via un appel à Visual Basic pour Applications (VBA). Ils ne peuvent pas être appelés à un autre moment. Par exemple, ils ne peuvent pas être appelés pendant les appels à la fonction [DllMain](https://docs.microsoft.com/windows/desktop/dlls/dllmain) ou à d’autres moments où le système d’exploitation a appelé la DLL, ou à partir d’un thread créé par la DLL. 
  
Les [fonctions Excel4v et Excel12v](excel4v-excel12v.md) acceptent leurs arguments en tant que tableau, tandis que les fonctions **Excel4** et **Excel12** acceptent leurs arguments en tant que liste de longueur variable dans la pile. À tous les autres égards, **Excel4** se comporte de la même manière **qu’Excel4v** et **Excel12** se comporte de la même manière **qu’Excel12v**.
  
```cs
int Excel4(int iFunction, LPXLOPER pxRes, int iCount, LPXLOPER argument1, ...);
int Excel12(int iFunction, LPXLOPER12 pxRes, int iCount, LPXLOPER12 argument1, ...);
```

## <a name="parameters"></a>Parameters

 _iFunction_ (**int**)
  
Nombre qui indique la commande, la fonction ou la fonction spéciale que vous voulez appeler. Pour obtenir la liste des valeurs  _iFunction_ valides, consultez la section Remarques suivante. 
  
 _pxRes_ (**LPXLOPER** ou **LPXLOPER12**)
  
Pointeur vers une **XLOPER** (avec **Excel4**) ou une **XLOPER12** (avec **Excel12)** qui conservera le résultat de la fonction évaluée.
  
 _iCount_ (**int**)
  
Nombre d’arguments ultérieurs qui seront transmis à la fonction. Dans les versions Excel jusqu’en 2003, ce nombre peut être n’importe quel nombre de 0 à 30. À compter Excel 2007, ce nombre peut être n’importe quel nombre entre 0 et 255.
  
 _argument1, ..._ (**LPXLOPER** ou **LPXLOPER12**)
  
Arguments facultatifs de la fonction. Tous les arguments doivent être des pointeurs vers **des valeurs XLOPER** ou **XLOPER12.** 
  
## <a name="return-value"></a>Valeur renvoyée

Renvoie l’une des valeurs d’un nombre **integer (int)** suivantes.
  
|**Valeur**|**Code de retour**|**Description**|
|:-----|:-----|:-----|
|0  <br/> |**xlretSuccess** <br/> |La fonction a été appelée avec succès. Cela ne signifie pas que la fonction n’a pas Excel d’erreur ; Pour le savoir, vous devez examiner le type et la valeur du paramètre _pxRes_ résultant.  <br/> |
|1  <br/> |**xlretAbort** <br/> |La commande ou la fonction a été interrompue anormalement (abandon interne). Cela peut se produire si une feuille macro XLM se ferme en appelant **CLOSE** ou si Excel est en mémoire. Si Excel renvoie cette erreur, la fonction d’appel doit se quitter immédiatement. La DLL est autorisée à appeler **xlFree** uniquement avant de quitter. Tous les autres appels à l’API C ne sont pas autorisés. L’utilisateur peut enregistrer n’importe quel travail de manière interactive à l’aide de la commande **Enregistrer** **dans** le menu Fichier.  <br/> |
|2  <br/> |**xlretInvXlfn** <br/> |Un numéro de fonction non valide a été fourni. Si vous utilisez des constantes du fichier d’en-tête Xlcall.h, cela ne doit pas se produire, sauf si vous appelez un appel qui n’est pas pris en charge dans la version de Excel que vous exécutez.  <br/> |
|4   <br/> |**xlretInvCount** <br/> |Un nombre d’arguments non valide a été entré. Dans les versions jusqu Excel 2003, le nombre maximal d’arguments qu’une fonction peut prendre est de 30. À compter Excel 2007, le nombre maximal est 255. Certains nécessitent un nombre fixe ou minimal d’arguments.  <br/> |
|8   <br/> |**xlretInvXloper** <br/> |Une **XLOPER ou** **XLOPER12** non valide a été transmise à la fonction ou un argument de type erroné a été utilisé.  <br/> |
|16   <br/> |**xlretStackOvfl** <br/> |Un dépassement de la pile s’est produit. Utilisez **xlStack** pour surveiller la quantité d’espace laissé sur la pile. Évitez d’allouer de très grandes matrices et structures locales (automatiques) sur la pile dans la mesure du possible ; les rendre statiques; (Notez qu’un dépassement de la pile peut se produire sans être détecté.)  <br/> |
|32  <br/> |**xlretFailed** <br/> |Échec d’une fonction équivalente à une commande. Cela équivaut à une commande de macro affichant la boîte de dialogue d’alerte d’erreur de macro.  <br/> |
|64  <br/> |**xlretUncalced** <br/> |Une tentative de déréférencement d’une cellule qui n’a pas encore été calculée a été tentée, car elle est programmée pour être recalculée après la cellule actuelle. Dans ce cas, la DLL doit retourner le contrôle Excel immédiatement. La DLL est autorisée à appeler **xlFree** uniquement avant de quitter. Tous les autres appels à l’API C ne sont pas autorisés. Pour plus d’informations sur les fonctions qui peuvent ou ne peuvent pas accéder aux valeurs des cellules qui n’ont pas été recalculées, voir Excel [Commands, Functions, and States](excel-commands-functions-and-states.md).  <br/> |
|128  <br/> |**xlretNotThreadSafe** <br/> |Une tentative d’appel d’une fonction qui n’est pas thread-safe ou qui peut ne pas l’être pendant un recalcul multithread du workbook a été tentée.  <br/> À compter Excel 2007, cette valeur est renvoyée et uniquement dans les fonctions de feuille de calcul XLL déclarées comme thread-safe.  <br/> |
|256  <br/> |**xlRetInvAsynchronousContext** <br/> |Le handle de fonction asynchrone n’est pas valide.  <br/> Cette valeur est utilisée uniquement par Excel 2010.  <br/> |
|512  <br/> |**xlRetNotClusterSafe** <br/> |L’appel n’est pas pris en charge sur les clusters.  <br/> Cette valeur est utilisée uniquement par Excel 2010.  <br/> |
   
## <a name="remarks"></a>Remarques

### <a name="valid-ifunction-values"></a>Valeurs iFunction valides

Les **valeurs iFunction** valides sont l’une des constantes **xlf...** ou **xlc...** définies dans le fichier d’en-tête Xlcall.h ou l’une des fonctions spéciales suivantes. 
  
|||||
|:-----|:-----|:-----|:-----|
|**xlAbort** <br/> |**xlEnableXLMsgs** <br/> |**xlGetInst** <br/> |**xlSheetNm** <br/> |
|**xlCoerce** <br/> |**xlFree** <br/> |**xlGetName** <br/> |**xlStack** <br/> |
|**xlDefineBinaryName** <br/> |**xlGetBinaryName** <br/> |**xlSet** <br/> |**xlUDF** <br/> |
|**xlDisableXLMsgs** <br/> |**xlGetHwnd** <br/> |**xlSheetId** <br/> ||
   
### <a name="different-types-of-functions"></a>Différents types de fonctions

 **Excel4** et **Excel12** font la distinction entre trois classes de fonctions. Les fonctions sont classées en fonction des trois états dans lesquels Excel peuvent appeler la DLL. 
  
- La classe 1 s’applique lorsque la DLL est appelée à partir d’une feuille de calcul suite à un recalcul. 
    
- La classe 2 s’applique lorsque la DLL est appelée à partir d’une macro de fonction ou d’une feuille de calcul où elle a été inscrite avec un signe de numéro (#) dans le texte du type.
    
- La classe 3 s’applique lorsqu’une DLL est appelée à partir d’un objet, d’une macro, d’un menu, d’une barre d’outils, d’une touche de raccourci, d’une méthode **ExecuteExcel4Macro** ou de la commande **Outils/Macro/Exécuter.** Pour plus d’informations, [voir Excel Commands, Functions et States](excel-commands-functions-and-states.md).
    
Le tableau suivant indique les fonctions valides dans chaque classe.
  
|**Classe 1**|**Classe 2**|**Classe 3**|
|:-----|:-----|:-----|
|N’importe quelle fonction de feuille de calcul  <br/> Toute fonction XLL uniquement **xl...** à l’exception **de xlSet**.  <br/> **xlfCaller** <br/> |N’importe quelle fonction de feuille de calcul  <br/> Toute **fonction xl...** à l’exception **de xlSet**.  <br/> Fonctions de feuille macro, y compris **xlfCaller,** qui retournent une valeur mais n’effectuent aucune action qui affecte l’espace de travail ou tout classeur ouvert.  <br/> |N’importe quelle fonction, **y compris xlSet** et les fonctions équivalentes à la commande.  <br/> |
   
### <a name="displaying-the-dialog-box-for-a-command-equivalent-function"></a>Affichage de la boîte de dialogue pour une Command-Equivalent de boîte de dialogue

Si une fonction équivalente à une commande est associée à une boîte de dialogue, vous pouvez définir le bit **xlPrompt** dans **iFunction**. Cela signifie que Excel affiche la boîte de dialogue appropriée avant d’exécuter la commande.
  
### <a name="writing-international-dlls"></a>Écriture de DLLs internationales

Si vous définissez le bit **xlIntl** dans **iFunction,** la fonction ou la commande est exécutée comme si elle était appelée à partir d’une feuille macro internationale. Cela signifie que la commande se comporte comme elle le ferait sur la version américaine de Excel, même si elle est en cours d’exécution sur une version internationale (localisée).
  
### <a name="xlretuncalced-or-xlretabort"></a>xlretUncalced ou xlretAbort

Après avoir reçu l’une de ces valeurs de retour, votre DLL doit nettoyer et retourner le contrôle Excel immédiatement. Les rappels dans Excel via l’API C, à l’exception de **xlFree,** sont désactivés après la réception de l’une de ces valeurs de retour.
  
## <a name="example"></a>Exemple

L’exemple suivant utilise la **fonction Excel12** pour sélectionner la cellule à partir de laquelle elle a été appelée. 
  
Cet exemple de code fait partie d’un exemple plus vaste fourni dans le SDK XLL Excel 2010, à l’emplacement suivant où vous avez installé le SDK :
  
\Samples\Example\Example.c.
  
> [!NOTE]
> Cette fonction appelle une macro de commande (xlcSelect) et, par conséquent, ne fonctionne que si elle est appelée à partir d’une feuille macro XLM. 
  
```cs
short WINAPI Excel12Example(void)
{
    XLOPER12 xRes;
    Excel12(xlfCaller, &xRes, 0);
    Excel12(xlcSelect, 0, 1, (LPXLOPER12)&xRes);
    Excel12(xlFree, 0, 1, (LPXLOPER12)&xRes);
    return 1;
}
```

## <a name="see-also"></a>Voir aussi



[Excel4v/Excel12v](excel4v-excel12v.md)

