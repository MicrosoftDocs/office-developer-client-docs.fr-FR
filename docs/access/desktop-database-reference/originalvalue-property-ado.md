---
title: OriginalValue, propriété (ADO)
TOCTitle: OriginalValue property (ADO)
ms:assetid: 02ffc728-4692-d439-e2a6-2f02cca53a71
ms:mtpsurl: https://msdn.microsoft.com/library/JJ248798(v=office.15)
ms:contentKeyID: 48542974
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
ms.openlocfilehash: 0724320e1aaa1e7bfd3ceab8cf54afd5921c7425
ms.sourcegitcommit: 8fe462c32b91c87911942c188f3445e85a54137c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2019
ms.locfileid: "32288181"
---
# <a name="originalvalue-property-ado"></a>OriginalValue, propriété (ADO)

**S’applique à** : Access 2013, Office 2013

Indique la valeur d'un objet [Field](field-object-ado.md) existant dans l'enregistrement avant qu'aucune modification n'ait été apportée.

## <a name="return-value"></a>Valeur renvoyée

Renvoie une valeur **Variant** représentant la valeur d'un champ avant modification.

## <a name="remarks"></a>Remarques

Utilisez la propriété **OriginalValue** pour renvoyer la valeur d'origine d'un champ depuis l'enregistrement actuel.

En *mode* de mise à jour immédiate (dans lequel le fournisseur écrit les modifications dans la source de données sous-jacente après l’appel de la méthode [Update),](update-method-ado.md) la propriété **OriginalValue** renvoie la valeur de champ qui existait avant toute modification (c’est-à-dire, depuis le dernier appel de méthode **Update).** C’est la même valeur que la méthode [CancelUpdate](cancelupdate-method-ado.md) utilise pour remplacer la propriété [Value](value-property-ado.md).

En *mode de mise à jour par lot* (dans lequel le fournisseur met en cache plusieurs modifications et ne les inscrit dans la source de données sous-jacente que lorsque vous appelez la méthode [UpdateBatch](updatebatch-method-ado.md)), la propriété **OriginalValue** renvoie la valeur du champ avant modification (c’est-à-dire avant le dernier appel de la méthode **UpdateBatch**). C’est cette valeur que la méthode [CancelBatch](cancelbatch-method-ado.md) utilise pour remplacer la propriété **Value**. Lorsque vous utilisez cette propriété en association avec la propriété [UnderlyingValue](underlyingvalue-property-ado.md), vous pouvez résoudre les conflits générés par les mises à jour par lot.

## <a name="record"></a>Record

Pour les objets [Record](record-object-ado.md), la propriété **OriginalValue** est vide pour les champs ajoutés avant l’appel de la méthode [Update](update-method-ado.md).

