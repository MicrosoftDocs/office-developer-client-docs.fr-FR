---
title: URL absolues et relatives
TOCTitle: Absolute and Relative URLs
ms:assetid: 79a1f793-7154-1c13-7dfe-a1b8cd64e1ea
ms:mtpsurl: https://msdn.microsoft.com/library/JJ249501(v=office.15)
ms:contentKeyID: 48545774
ms.date: 09/18/2015
mtps_version: v=office.15
ms.openlocfilehash: 6bc0cb3086f8fdfa032c005f7e2d219dab56999b
ms.sourcegitcommit: 19aca09c5812cfb98b68b5d4604dcaa814479df7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2018
ms.locfileid: "25471845"
---
# <a name="absolute-and-relative-urls"></a><span data-ttu-id="9b5ab-102">URL absolues et relatives</span><span class="sxs-lookup"><span data-stu-id="9b5ab-102">Absolute and Relative URLs</span></span>

<span data-ttu-id="9b5ab-103">**S’applique à**: Access 2013 | Office 2013</span><span class="sxs-lookup"><span data-stu-id="9b5ab-103">**Applies to**: Access 2013 | Office 2013</span></span> 

<span data-ttu-id="9b5ab-p101">Une URL spécifie l'emplacement d'une cible sur un ordinateur local ou en réseau, par exemple un fichier, un répertoire, une page HTML, une image, un programme, etc *.* Dans cette présentation, une *URL absolue* a la forme suivante :</span><span class="sxs-lookup"><span data-stu-id="9b5ab-p101">A URL specifies the location of a target stored on a local or networked computer, such as a file, directory, HTML page, image, program, and so on *.* In this discussion, an *absolute URL* is of the form:</span></span>

<span data-ttu-id="9b5ab-106">*modèle://serveur/chemin d'accès/ressource*</span><span class="sxs-lookup"><span data-stu-id="9b5ab-106">*scheme://server/path/resource*</span></span>

<span data-ttu-id="9b5ab-107">où :</span><span class="sxs-lookup"><span data-stu-id="9b5ab-107">where:</span></span>

  - <span data-ttu-id="9b5ab-108">*modèle*</span><span class="sxs-lookup"><span data-stu-id="9b5ab-108">*scheme*</span></span>

  - <span data-ttu-id="9b5ab-109">Spécifie comment la *ressource* est accessible.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-109">Specifies how the *resource* is to be accessed.</span></span>

  - <span data-ttu-id="9b5ab-110">*serveur*</span><span class="sxs-lookup"><span data-stu-id="9b5ab-110">*server*</span></span>

  - <span data-ttu-id="9b5ab-111">Spécifie le nom de l’ordinateur où se trouve la *ressource* .</span><span class="sxs-lookup"><span data-stu-id="9b5ab-111">Specifies the name of the computer where the *resource* is located.</span></span>

  - <span data-ttu-id="9b5ab-112">*chemin d'accès*</span><span class="sxs-lookup"><span data-stu-id="9b5ab-112">*path*</span></span>

  - <span data-ttu-id="9b5ab-p102">Indique la séquence de répertoires menant à la cible. Si la *ressource* est omise, la cible est le dernier répertoire figurant dans le *chemin d'accès*.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-p102">Specifies the sequence of directories leading to the target. If *resource* is omitted, the target is the last directory in *path*.</span></span>

  - <span data-ttu-id="9b5ab-115">*ressource*</span><span class="sxs-lookup"><span data-stu-id="9b5ab-115">*resource*</span></span>

  - <span data-ttu-id="9b5ab-116">S’il est inclus, la *ressource* est la cible et est généralement le nom d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-116">If included, *resource* is the target, and is typically the name of a file.</span></span> <span data-ttu-id="9b5ab-117">Il peut être un *fichier simple* contenant un flux binaire d’octets ou un *document structuré,* contenant une ou plusieurs stockages et flux binaires d’octets.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-117">It may be a *simple file,* containing a single binary stream of bytes, or a *structured document,* containing one or more storages and binary streams of bytes.</span></span>

<span data-ttu-id="9b5ab-118">Une *URL absolue* contient toutes les informations nécessaires pour localiser une ressource.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-118">An *absolute URL* contains all the information necessary to locate a resource.</span></span>

<span data-ttu-id="9b5ab-p104">Une *URL relative* localise une ressource en utilisant une URL absolue comme point de départ. Dans la pratique, l' « URL complète » de la cible est spécifiée par la concaténation des URL absolue et relative. Une URL relative n'est constituée en général que du *chemin d'accès* et, éventuellement, de la *ressource*, mais pas du *modèle* ni du *serveur*.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-p104">A *relative URL* locates a resource using an absolute URL as a starting point. In effect, the "complete URL" of the target is specified by concatenating the absolute and relative URLs. A relative URL typically consists only of the *path*, and optionally, the *resource*, but no *scheme* or *server*.</span></span>

## <a name="url-scheme-registration"></a><span data-ttu-id="9b5ab-122">Enregistrement du modèle d'URL</span><span class="sxs-lookup"><span data-stu-id="9b5ab-122">URL Scheme Registration</span></span>

<span data-ttu-id="9b5ab-123">Si un fournisseur prend en charge plusieurs URL, il sera enregistré pour un ou plusieurs modèles d'URL.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-123">If a provider supports URLs, it will register for one or more URL schemes.</span></span> <span data-ttu-id="9b5ab-124">C'est-à-dire que les URL utilisant ces modèles appelleront automatiquement le fournisseur enregistré.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-124">This means that any URLs using this scheme will automatically invoke the registered provider.</span></span> <span data-ttu-id="9b5ab-125">Par exemple, le schéma *http* est enregistré pour le [Fournisseur Microsoft OLE DB pour la publication Internet](microsoft-ole-db-provider-for-internet-publishing.md).</span><span class="sxs-lookup"><span data-stu-id="9b5ab-125">For example, the *http* scheme is registered to the [Microsoft OLE DB Provider for Internet Publishing](microsoft-ole-db-provider-for-internet-publishing.md).</span></span> <span data-ttu-id="9b5ab-126">ADO suppose que toutes les URL avec le préfixe « http » représentent des dossiers ou fichiers Web à utiliser avec le fournisseur de publication Internet.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-126">ADO assumes all URLs prefixed with "http" represent Web folders or files to be used with the Internet Publishing Provider.</span></span> <span data-ttu-id="9b5ab-127">Pour plus d'informations sur les modèles enregistrés par votre fournisseur, consultez la documentation de ce dernier.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-127">For information about the schemes registered by your provider, see your provider documentation.</span></span>

## <a name="defining-context-with-a-url"></a><span data-ttu-id="9b5ab-128">Définition d'un contexte avec une URL</span><span class="sxs-lookup"><span data-stu-id="9b5ab-128">Defining Context with a URL</span></span>

<span data-ttu-id="9b5ab-p106">Une des fonctions d'une connexion ouverte, représentée par un objet [Connection](connection-object-ado.md), est de limiter les opérations ultérieures sur la source de données représentée par cette connexion. En d'autres termes, la connexion définit le contexte des opérations suivantes.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-p106">One function of an open connection, represented by a [Connection](connection-object-ado.md) object, is to restrict subsequent operations to the data source represented by that connection. That is, the connection defines the context for subsequent operations.</span></span>

<span data-ttu-id="9b5ab-p107">Avec ADO 2.5, une URL absolue peut également définir un contexte. Par exemple, lorsqu'un objet [Record](record-object-ado.md) est ouvert avec une URL absolue, un objet **Connection** est implicitement créé pour représenter la ressource spécifiée par l'URL.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-p107">With ADO 2.5, an absolute URL may also define a context. For example, when a [Record](record-object-ado.md) object is opened with an absolute URL, a **Connection** object is implicitly created to represent the resource specified by the URL.</span></span>

<span data-ttu-id="9b5ab-133">Une URL absolue qui définit un contexte peut être spécifiée dans le paramètre *ActiveConnection* de la méthode [Open](open-method-ado-record.md) de l’objet **Record** .</span><span class="sxs-lookup"><span data-stu-id="9b5ab-133">An absolute URL that defines a context may be specified in the *ActiveConnection* parameter of the **Record** object [Open](open-method-ado-record.md) method.</span></span> <span data-ttu-id="9b5ab-134">Vous pouvez également spécifier une URL absolue comme valeur de la nouvelle « URL**=**« mot clé dans le paramètre *ConnectionString* méthode [Open](open-method-ado-connection.md) de **connexion** objet et la méthode [Open](open-method-ado-recordset.md) \*ActiveConnection de l’objet [Recordset](recordset-object-ado.md) \*paramètre.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-134">An absolute URL may also be specified as the value of the new "URL**=**" keyword in the **Connection** object [Open](open-method-ado-connection.md) method *ConnectionString* parameter, and the [Recordset](recordset-object-ado.md) object [Open](open-method-ado-recordset.md) method *ActiveConnection* parameter.</span></span>

<span data-ttu-id="9b5ab-135">Le contexte peut également être défini avec un objet **Record** ou **Recordset** ouvert qui représente un répertoire, car ces objets disposent déjà d'un objet **Connection** implicitement ou explicitement déclaré qui spécifie le contexte.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-135">Context may also be defined with an open **Record** or **Recordset** object that represents a directory because these objects already have an implicitly or explicitly declared **Connection** object that specifies context.</span></span>

## <a name="scoped-operations"></a><span data-ttu-id="9b5ab-136">Opérations de portée définie</span><span class="sxs-lookup"><span data-stu-id="9b5ab-136">Scoped Operations</span></span>

<span data-ttu-id="9b5ab-137">Le contexte définit simultanément une *portée* — autrement dit, le répertoire et ses sous-répertoires pouvant participer aux opérations suivantes.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-137">The context simultaneously defines a *scope* — that is, the directory and its subdirectories that may participate in subsequent operations.</span></span> <span data-ttu-id="9b5ab-138">L'objet **Record** comporte plusieurs méthodes de portée définie, notamment [CopyRecord](copyrecord-method-ado.md), [MoveRecord](moverecord-method-ado.md) et [DeleteRecord](https://msdn.microsoft.com/library/jj249832\(v=office.15\)) qui s'appliquent à un répertoire et à ses sous-répertoires.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-138">The **Record** object has several scoped methods, including [CopyRecord](copyrecord-method-ado.md), [MoveRecord](moverecord-method-ado.md), and [DeleteRecord](https://msdn.microsoft.com/library/jj249832\(v=office.15\)), that operate on a directory and all its subdirectories.</span></span>

## <a name="relative-urls-as-command-text"></a><span data-ttu-id="9b5ab-139">URL relatives sous la forme de texte de commande</span><span class="sxs-lookup"><span data-stu-id="9b5ab-139">Relative URLs as Command Text</span></span>

<span data-ttu-id="9b5ab-140">Chaîne spécifiant une commande à exécuter sur la source de données peut être spécifiée dans le paramètre *CommandText* de méthode **Connection** objet **Execute** et le paramètre de la *Source* de méthode **Open** **Recordset** objet.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-140">A string specifying a command to be executed on the data source may be specified in the **Connection** object **Execute** method *CommandText* parameter, and the **Recordset** object **Open** method *Source* parameter.</span></span>

<span data-ttu-id="9b5ab-141">Une URL relative peut être spécifiée dans le paramètre *CommandText* ou *Source* .</span><span class="sxs-lookup"><span data-stu-id="9b5ab-141">A relative URL may be specified in the *CommandText* or *Source* parameter.</span></span> <span data-ttu-id="9b5ab-142">L'URL relative ne spécifie pas une commande (comme une commande SQL) ; elle est simplement spécifiée dans ces paramètres.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-142">The relative URL does not actually specify a command (such as an SQL command); it is merely specified in those parameters.</span></span> <span data-ttu-id="9b5ab-143">En outre, le contexte de la connexion active doit être une URL absolue et le paramètre *Option* doit être défini **adCmdTableDirect**.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-143">In addition, the context of the active connection must be an absolute URL, and the *Option* parameter must be set to **adCmdTableDirect**.</span></span>

<span data-ttu-id="9b5ab-144">Par exemple, un objet **Recordset** peut être ouvert sur le fichier Readme25.txt du répertoire Winnt/system32 comme suit :</span><span class="sxs-lookup"><span data-stu-id="9b5ab-144">For example, a **Recordset** could be opened on the Readme25.txt file of the Winnt/system32 directory like this:</span></span>

```vb
recordset.Open "system32/Readme25.txt", "URL=https://YourServer/Winnt/",,,adCmdTableDirect 
```

<span data-ttu-id="9b5ab-145">L’URL absolue dans la chaîne de connexion spécifie le serveur (YourServer) et le chemin d’accès () et le chemin d’accès (Winnt).</span><span class="sxs-lookup"><span data-stu-id="9b5ab-145">The absolute URL in the connection string specifies the server (YourServer ) and the path () and the path (Winnt ).</span></span> <span data-ttu-id="9b5ab-146">Cette URL définit également le contexte.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-146">This URL also defines the context.</span></span>

<span data-ttu-id="9b5ab-147">L’URL relative dans le texte de commande utilise l’URL absolue comme point de départ et spécifie le reste du chemin d’accès (system32) et le fichier à ouvrir () et le fichier à ouvrir (Readme25.txt).</span><span class="sxs-lookup"><span data-stu-id="9b5ab-147">The relative URL in the command text uses the absolute URL as a starting point and specifies the remainder of the path (system32 ) and the file to open () and the file to open (Readme25.txt ).</span></span>

<span data-ttu-id="9b5ab-148">Le champ d’options () indique que le type de commande est une URL relative.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-148">The options field () indicates that the command type is a relative URL.</span></span>

<span data-ttu-id="9b5ab-149">Autre exemple, le code suivant ouvre un **objet Recordset** sur le contenu du répertoire :</span><span class="sxs-lookup"><span data-stu-id="9b5ab-149">As another example, the following code will open a **Recordset** on the contents of the directory:</span></span>

```vb
recordset.Open "", "URL=https://YourServer/Winnt/",,,adCmdTableDirect 
```

## <a name="ole-db-provider-supplied-url-schemes"></a><span data-ttu-id="9b5ab-150">Modèles d'URL fournis par le fournisseur OLE DB</span><span class="sxs-lookup"><span data-stu-id="9b5ab-150">OLE DB Provider-Supplied URL Schemes</span></span>

<span data-ttu-id="9b5ab-151">La première partie d’une URL complète est le *schéma* utilisé pour accéder à la ressource identifiée par le reste de l’URL.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-151">The leading part of a fully-qualified URL is the *scheme* used to access the resource identified by the remainder of the URL.</span></span> <span data-ttu-id="9b5ab-152">Par exemple, HTTP (HyperText Transfer Protocol) et FTP (File Transfer Protocol).</span><span class="sxs-lookup"><span data-stu-id="9b5ab-152">Examples are HTTP (HyperText Transfer Protocol) and FTP (File Transfer Protocol).</span></span>

<span data-ttu-id="9b5ab-153">ADO prend en charge les fournisseurs OLE DB qui reconnaissent leurs propres schémas d’URL.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-153">ADO supports OLE DB providers that recognize their own URL schemes.</span></span> <span data-ttu-id="9b5ab-154">Par exemple, le [Fournisseur Microsoft OLE DB pour la publication Internet](microsoft-ole-db-provider-for-internet-publishing.md)*,* qui accède aux fichiers Windows 2000 « publiés », reconnaît le modèle HTTP existant.</span><span class="sxs-lookup"><span data-stu-id="9b5ab-154">For example, the [Microsoft OLE DB Provider for Internet Publishing](microsoft-ole-db-provider-for-internet-publishing.md)*,* which accesses "published" Windows 2000 files, recognizes the existing HTTP scheme.</span></span>
