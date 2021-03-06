---
title: Élément SnapSettings (DocumentSettings_Type complexType) (Visio XML)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: 6e86e943-bd29-0a7b-3d6a-d91281f98777
description: Spécifie les objets sur qui les formes s’ancrent lorsque l’a snap est actif dans la fenêtre.
ms.openlocfilehash: 8d4be35a4cd66a1d3914dbda8162f4acb3d05bfa
ms.sourcegitcommit: e7b38e37a9d79becfd679e10420a19890165606d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/29/2019
ms.locfileid: "34540293"
---
# <a name="snapsettings-element-documentsettings_type-complextype-visio-xml"></a>Élément SnapSettings (DocumentSettings_Type complexType) (Visio XML)

Spécifie les objets sur qui les formes s’ancrent lorsque l’a snap est actif dans la fenêtre.
  
## <a name="element-information"></a>Informations sur l’élément

|||
|:-----|:-----|
|**Type d’élément** <br/> |[SnapSettings_Type](snapsettings_type-complextypevisio-xml.md) <br/> |
|**Namespace** <br/> |http://schemas.microsoft.com/office/visio/2012/main  <br/> |
|**Fichier de schéma** <br/> |VisioSchema15.xsd  <br/> |
|**Composants de document** <br/> |document.xml  <br/> |
   
## <a name="definition"></a>Définition

```XML
< xs:element name="SnapSettings" type="SnapSettings_Type" minOccurs="0" maxOccurs="1" >
</xs:element >
```

## <a name="elements-and-attributes"></a>Éléments et attributs

Si le schéma définit des exigences spécifiques, telles que **séquence**, **minOccurs**, **maxOccurs** et **choix**, voir la section de définition. 
  
### <a name="parent-elements"></a>Éléments parents

|**Élément**|**Type (Type)**|**Description**|
|:-----|:-----|:-----|
|[DocumentSettings](documentsettings-element-visiodocument_type-complextypevisio-xml.md) <br/> |[DocumentSettings_Type](documentsettings_type-complextypevisio-xml.md) <br/> |Contient des éléments qui spécifient les paramètres de document.  <br/> |
   
### <a name="child-elements"></a>Éléments enfants

Aucun.
  
### <a name="attributes"></a>Attributs

Aucun.
  
## <a name="remarks"></a>Remarques

La valeur peut être une somme des valeurs du tableau suivant.
  
|**Valeur**|**Description**|
|:-----|:-----|
|0  <br/> |Aligne sur rien.  <br/> |
|1  <br/> |Ancrer sous-sections de règle.  <br/> |
|2  <br/> |Ancrer grille.  <br/> |
|4   <br/> |Aligne sur les repères.  <br/> |
|8   <br/> |Aligne sur les poignées de sélection.  <br/> |
|16   <br/> |Aligne sur les sommets.  <br/> |
|32  <br/> |Aligne sur les points de connexion.  <br/> |
|256  <br/> |Ancrer bords visibles des formes.  <br/> |
|512  <br/> |Ancrer’alignement.  <br/> |
|1024  <br/> |Aligne sur les options d'extension de forme.  <br/> |
|32768  <br/> |Ancrer désactivé.  <br/> |
|65536  <br/> |Aligner sur les intersections.  <br/> |
   

