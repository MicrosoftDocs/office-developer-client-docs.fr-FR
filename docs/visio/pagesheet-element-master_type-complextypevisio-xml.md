---
title: Élément PageSheet (Master_Type complexType) (Visio XML)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: 824fbeb0-1a2f-35a0-50e3-c57143dc21ab
description: Spécifie les propriétés de la page de dessin associée à la page de dessin.
ms.openlocfilehash: 94fde64b130c2a05c4bd70c97552fe4218171ce7
ms.sourcegitcommit: e7b38e37a9d79becfd679e10420a19890165606d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/29/2019
ms.locfileid: "34540615"
---
# <a name="pagesheet-element-master_type-complextype-visio-xml"></a>Élément PageSheet (Master_Type complexType) (Visio XML)

Spécifie les propriétés de la page de dessin associée à la page de dessin.
  
## <a name="element-information"></a>Informations sur l’élément

|||
|:-----|:-----|
|**Type d’élément** <br/> |[PageSheet_Type](pagesheet_type-complextypevisio-xml.md) <br/> |
|**Namespace** <br/> |http://schemas.microsoft.com/office/visio/2012/main  <br/> |
|**Fichier de schéma** <br/> |VisioSchema15.xsd  <br/> |
|**Composants de document** <br/> |masters.xml  <br/> |
   
## <a name="definition"></a>Définition

```XML
< xs:element name="PageSheet" type="PageSheet_Type" minOccurs="0" maxOccurs="1" >
</xs:element >
```

## <a name="elements-and-attributes"></a>Éléments et attributs

Si le schéma définit des exigences spécifiques, telles que **séquence**, **minOccurs**, **maxOccurs** et **choix**, voir la section de définition. 
  
### <a name="parent-elements"></a>Éléments parents

|**Élément**|**Type (Type)**|**Description**|
|:-----|:-----|:-----|
|[Master](master-element-masters_type-complextypevisio-xml.md) <br/> |[Master_Type](master_type-complextypevisio-xml.md) <br/> |Spécifie une master dans un dessin.  <br/> |
   
### <a name="child-elements"></a>Éléments enfants

Aucun.
  
### <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|**Valeurs possibles**|
|:-----|:-----|:-----|:-----|:-----|
|FillStyle  <br/> |xsd:unsignedInt  <br/> |facultatif  <br/> |spécifie l’ID de la feuille de style à partir de laquelle hériter la mise en forme du remplissage. Il DOIT s’agit de la valeur de **l’attribut ID** associé à un **StyleSheet_Type** dans le dessin.  <br/> |Valeurs du type xsd:unsignedInt.  <br/> |
|LineStyle  <br/> |xsd:unsignedInt  <br/> |facultatif  <br/> |Spécifie l’ID de la feuille de style à partir de laquelle hériter la mise en forme des lignes. Il DOIT s’agit de la valeur de **l’attribut ID** associé à un **StyleSheet_Type** dans le dessin.  <br/> |Valeurs du type xsd:unsignedInt.  <br/> |
|TextStyle  <br/> |xsd:unsignedInt  <br/> |facultatif  <br/> |Spécifie l’ID de la feuille de style à partir de laquelle hériter la mise en forme du texte. Il DOIT s’agit de la valeur de **l’attribut ID** associé à un **StyleSheet_Type** dans le dessin.  <br/> |Valeurs du type xsd:unsignedInt.  <br/> |
|UniqueID  <br/> |xsd:string  <br/> |facultatif  <br/> |ID unique de l’élément au sein de son élément parent.  <br/> |Valeurs du type xsd:string.  <br/> |
   

