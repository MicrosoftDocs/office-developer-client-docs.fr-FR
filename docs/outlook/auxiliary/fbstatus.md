---
title: FBStatus
manager: soliver
ms.date: 11/16/2014
ms.audience: Developer
ms.topic: overview
localization_priority: Normal
ms.assetid: f2d6a11e-847d-6bbe-cd77-e78ee961cb12
description: Une éumération pour l’état de libre/occupé des blocs de libre/occupé.
ms.openlocfilehash: 2a08ef142f9baddd453166c0ebcb989e69a51ceb
ms.sourcegitcommit: 8657170d071f9bcf680aba50b9c07f2a4fb82283
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33424090"
---
# <a name="fbstatus"></a>FBStatus

Une éumération pour l’état de libre/occupé des blocs de libre/occupé.
  
## <a name="quick-info"></a>Informations rapides

```cpp
enum  
    { 
      fbFree      = 0, 
      fbTentative = fbFree + 1, 
      fbBusy      = fbTentative + 1, 
      fbOutOfOffice = fbBusy + 1 
    }

```

## <a name="remarks"></a>Remarques

L’état de libre/occupé d’un bloc de temps détermine la façon dont il s’affiche dans un calendrier : **Libre,** **Occupé,** **Provisoire** ou **Hors** Office . 
  
## <a name="see-also"></a>Voir aussi

- [FBBlock_1](fbblock_1.md)
- [IEnumFBBlock::Next](ienumfbblock-next.md)

