---
title: Élément RowKeyValue (PrimaryKey_Type complexType) (Visio XML)
manager: soliver
ms.date: 03/09/2015
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: 9077ad4b-c539-c0c8-d268-9a009990abdd
description: Spécifie la valeur d’une clé primaire pour une ligne individuelle d’un recordset.
ms.openlocfilehash: b21f479a9c0404dc8a3b737208e4d0d634b556f4
ms.sourcegitcommit: e7b38e37a9d79becfd679e10420a19890165606d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/29/2019
ms.locfileid: "34538129"
---
# <a name="rowkeyvalue-element-primarykey_type-complextype-visio-xml"></a>Élément RowKeyValue (PrimaryKey_Type complexType) (Visio XML)

Spécifie la valeur d’une clé primaire pour une ligne individuelle d’un recordset.
  
## <a name="element-information"></a>Informations sur l’élément

|||
|:-----|:-----|
|**Type d’élément** <br/> |[RowKeyValue_Type](rowkeyvalue_type-complextypevisio-xml.md) <br/> |
|**Namespace** <br/> |http://schemas.microsoft.com/office/visio/2012/main  <br/> |
|**Fichier de schéma** <br/> |VisioSchema15.xsd  <br/> |
|**Composants de document** <br/> |recordsets.xml  <br/> |
   
## <a name="definition"></a>Définition

```XML
< xs:element name="RowKeyValue" type="RowKeyValue_Type" minOccurs="0" maxOccurs="unbounded" >
</xs:element >
```

## <a name="elements-and-attributes"></a>Éléments et attributs

Si le schéma définit des exigences spécifiques, telles que **séquence**, **minOccurs**, **maxOccurs** et **choix**, voir la section de définition. 
  
### <a name="parent-elements"></a>Éléments parents

|**Élément**|**Type (Type)**|**Description**|
|:-----|:-----|:-----|
|[PrimaryKey](primarykey-element-datarecordset_type-complextypevisio-xml.md) <br/> |[PrimaryKey_Type](primarykey_type-complextypevisio-xml.md) <br/> |Spécifie une clé primaire d’un recordset.  <br/> |
   
### <a name="child-elements"></a>Éléments enfants

Aucun.
  
### <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|**Valeurs possibles**|
|:-----|:-----|:-----|:-----|:-----|
|RowID  <br/> |xsd:unsignedInt  <br/> |obligatoire  <br/> |Valeur unique qui identifie une ligne d’un recordset.  <br/> |Valeurs du type xsd:unsignedInt.  <br/> |
|Valeur  <br/> |xsd:string  <br/> |obligatoire  <br/> |Valeur de la clé primaire pour cette ligne du recordset.  <br/> |Valeurs du type xsd:string.  <br/> |
   

