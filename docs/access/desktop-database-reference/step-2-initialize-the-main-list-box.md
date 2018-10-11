---
title: 'Étape 2 : initialiser la zone de liste principale'
TOCTitle: 'Step 2: Initialize the Main List Box'
ms:assetid: 81e4dcfd-6ee0-b5f9-9ea3-026c38c26bf0
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249562(v=office.15)
ms:contentKeyID: 48545967
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: f2b6b4d79262796aa8e9090d673a439a550f042c
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471510"
---
# <a name="step-2-initialize-the-main-list-box"></a><span data-ttu-id="0fa0d-102">Étape 2 : initialiser la zone de liste principale</span><span class="sxs-lookup"><span data-stu-id="0fa0d-102">Step 2: Initialize the Main List Box</span></span>


<span data-ttu-id="0fa0d-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="0fa0d-103">**Applies to**: Access 2013 | Office 2013</span></span>

## <a name="step-2-initialize-the-main-list-box"></a><span data-ttu-id="0fa0d-104">Étape 2 : initialiser la zone de liste principale</span><span class="sxs-lookup"><span data-stu-id="0fa0d-104">Step 2: Initialize the Main List Box</span></span>

<span data-ttu-id="0fa0d-105">**Pour déclarer les objets Record et Recordset globaux**</span><span class="sxs-lookup"><span data-stu-id="0fa0d-105">**To declare global Record and Recordset objects**</span></span>

  - <span data-ttu-id="0fa0d-106">Insérez le code suivant dans (Général) (Déclarations) pour Form1 :</span><span class="sxs-lookup"><span data-stu-id="0fa0d-106">Insert the following code into the (General) (Declarations) for Form1:</span></span>
    
    ```vb 
     
    Option Explicit 
    Dim grec As Record 
    Dim grs As Recordset 
    ```
    
    <span data-ttu-id="0fa0d-107">Ce code déclare des références pour les objets globaux **Record** et **Recordset** qui seront utilisés par la suite dans ce scénario.</span><span class="sxs-lookup"><span data-stu-id="0fa0d-107">This code declares global object references for **Record** and **Recordset** objects that will be used later in this scenario.</span></span>

<span data-ttu-id="0fa0d-108">**Pour se connecter à une URL et remplir lstMain**</span><span class="sxs-lookup"><span data-stu-id="0fa0d-108">**To connect to a URL and populate lstMain**</span></span>

  - <span data-ttu-id="0fa0d-109">Insérez le code suivant dans le gestionnaire d'événements Form Load pour Form1 :</span><span class="sxs-lookup"><span data-stu-id="0fa0d-109">Insert the following code into the Form Load event handler for Form1:</span></span>
    
    ```vb 
     
    Private Sub Form_Load() 
        Set grec = New Record 
        Set grs = New Recordset 
        grec.Open "", "URL=https://servername/foldername/", , _ 
            adOpenIfExists Or adCreateCollection 
        Set grs = grec.GetChildren 
        While Not grs.EOF 
            lstMain.AddItem grs(0) 
            grs.MoveNext 
        Wend 
    End Sub 
    ```
    
    <span data-ttu-id="0fa0d-110">Ce code instancie les objets **Record** et **Recordset** globaux.</span><span class="sxs-lookup"><span data-stu-id="0fa0d-110">This code instantiates the global **Record** and **Recordset** objects.</span></span> <span data-ttu-id="0fa0d-111">L’objet **Record**, grec, est ouvert avec l’URL définie comme **ActiveConnection**.</span><span class="sxs-lookup"><span data-stu-id="0fa0d-111">The **Record**, grec, is opened with a URL specified as the **ActiveConnection**.</span></span> <span data-ttu-id="0fa0d-112">Si l'URL existe, elle est ouverte ; si elle n'existe pas encore, elle est créée.</span><span class="sxs-lookup"><span data-stu-id="0fa0d-112">If the URL exists, it is opened; if it does not already exist, it is created.</span></span> <span data-ttu-id="0fa0d-113">Notez que vous devez remplacer « https://servername/foldername/ » par une URL valide de votre environnement.</span><span class="sxs-lookup"><span data-stu-id="0fa0d-113">Note that you should replace "https://servername/foldername/" with a valid URL from your environment.</span></span> <span data-ttu-id="0fa0d-114">L' **objet Recordset**, grs, est ouvert sur les enfants de l’objet **Record**, grec.</span><span class="sxs-lookup"><span data-stu-id="0fa0d-114">The **Recordset**, grs, is opened on the children of the **Record**, grec.</span></span> <span data-ttu-id="0fa0d-115">lstMain est ensuite rempli avec les noms de fichier des ressources publiées au niveau de l'URL.</span><span class="sxs-lookup"><span data-stu-id="0fa0d-115">Then lstMain is populated with the file names of the resources published to the URL.</span></span>
