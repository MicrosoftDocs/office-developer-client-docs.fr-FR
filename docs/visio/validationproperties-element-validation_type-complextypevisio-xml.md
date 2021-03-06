---
title: Élément ValidationProperties (Validation_Type complexType) (Visio XML)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: a51a60c9-479b-7d7b-860f-bb46fc8b4d63
description: Encapsule les propriétés liées à la validation du document.
ms.openlocfilehash: 35e6f3f13eecef826fdef0d664bba35fceb0e069
ms.sourcegitcommit: e7b38e37a9d79becfd679e10420a19890165606d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/29/2019
ms.locfileid: "34538521"
---
# <a name="validationproperties-element-validation_type-complextype-visio-xml"></a>Élément ValidationProperties (Validation_Type complexType) (Visio XML)

Encapsule les propriétés liées à la validation du document.
  
## <a name="element-information"></a>Informations sur l’élément

|||
|:-----|:-----|
|**Type d’élément** <br/> |[ValidationProperties_Type](validationproperties_type-complextypevisio-xml.md) <br/> |
|**Namespace** <br/> |http://schemas.microsoft.com/office/visio/2012/main  <br/> |
|**Fichier de schéma** <br/> |VisioSchema15.xsd  <br/> |
|**Composants de document** <br/> |validation.xml  <br/> |
   
## <a name="definition"></a>Définition

```XML
< xs:element name="ValidationProperties" type="ValidationProperties_Type" minOccurs="0" maxOccurs="1" >
</xs:element >
```

## <a name="elements-and-attributes"></a>Éléments et attributs

Si le schéma définit des exigences spécifiques, telles que **séquence**, **minOccurs**, **maxOccurs** et **choix**, voir la section de définition. 
  
### <a name="parent-elements"></a>Éléments parents

|**Élément**|**Type (Type)**|**Description**|
|:-----|:-----|:-----|
|[Valider](validation-elementvisio-xml.md) <br/> |[Validation_Type](validation_type-complextypevisio-xml.md) <br/> |Stocke les informations relatives à la validation du diagramme pour le document.  <br/> |
   
### <a name="child-elements"></a>Éléments enfants

Aucun.
  
### <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|**Valeurs possibles**|
|:-----|:-----|:-----|:-----|:-----|
|LastValidated  <br/> |xsd:dateTime  <br/> |obligatoire  <br/> |Date et heure de la dernière validation du document.  <br/> |Valeurs du type xsd:dateTime.  <br/> |
|ShowIgnored  <br/> |xsd:boolean  <br/> |obligatoire  <br/> |Spécifie s’il faut afficher les problèmes de validation ignorés dans la fenêtre Problèmes.  <br/> |Valeurs du type xsd:boolean.  <br/> |
   

