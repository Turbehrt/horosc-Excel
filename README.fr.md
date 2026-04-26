# Horosc pour Excel

[![en](https://img.shields.io/badge/lang-en-red.svg)](/README.md)
[![fr](https://img.shields.io/badge/lang-fr-blue.svg)](/README.fr.md)

Version 1 (2026)

## Contexte du projet
L'ouvrage de John D. North, *Horoscopes and History*, Londres : Warburg Institute, 1986, publiait en annexe (appendix 4 p. 197-218) le code d’un programme en Pascal pour MS-DOS, intitulé **HOROSC** et visant à calculer et contrôler la domification d’un horoscope en suivant les 7 principales méthodes historiques.

En raison de l’obsolescence du langage Pascal, ce programme devenait très difficilement accessible. Initié en 2021, le projet _[Horosc for Google Sheets](https://github.com/Turbehrt/horosc-GoogleSheets)_ portait le code en Google Script pour un usage dans Google Sheets.

La façon dont fonctionne Google Apps Script favorise l'exécution de méthodes complètes (à la manière du programme Horosc inital) aux dépens d'un usage plus flexible de fonctions intermédiaires. Le présent projet, lancé en 2025, porte le code en VBA pour Microsoft Excel, offrant une plus grande variété de fonctions plus aisément combinables.

> [!TIP]
> Depuis la version 2 de _[Horosc for Google Sheets](https://github.com/Turbehrt/horosc-GoogleSheets)_, les deux applications utilisent les mêmes noms de fonctions et d'arguments, de façon à faciliter une utilisation parallèle. Il est généralement plus rapide d'utiliser la feuille de calcul Excel sur un ordinateur local que de lancer les scripts Google Apps dans le cloud, notamment en cas d'utilisation de plusieurs fonctions intermédiaires. Néanmoins, cela nécessite d'activer les macros en local.

## Principe général
Le programme écrit par John D. North proposait deux approches :

### Méthode A
En connaissant :
*	la valeur de l’obliquité de l’écliptique,
*	la latitude géographique du lieu d’observation,
*	et la longitude écliptique de l’Ascendant

le programe fournit :
* le calcul théorique des longitudes et/ou ascensions droites sexagésimales des 6 premières maisons (les 6 suivantes sont induites par symétrie) selon les 7 méthodes de domification historique

Cette méthode est restituée ici par la formule matricielle `computeLongitudesAllMethodsLatitude(obliquité de l’écliptique, latitude géographique, longitude de l’ascendant, nombre de lignes)`.
Elle produit deux tableaux, respectivement pour les longitudes et les ascensions droites, séparés par le nombre de lignes renseigné.

### Méthode B
En connaissant :
*	la valeur de l'obliquité de l'écliptique,
*	les longitudes des 6 premières maisons (transcrites d'une source historique),
*	et une marge d’erreur ou d’arrondi

le programe fournit :
*	un calcul théorique de la latitude du lieu d’observation (avec un intervalle correspondant à la marge d’erreur, appliquée à l’ascension droite de l’ascendant ou du milieu du ciel)
*	une comparaison avec les longitudes théoriques (calculées en considérant exacts l'ascendant et le milieu du ciel) selon les 7 méthodes de domification historique, avec coefficient de qualité (permettant généralement d’identifier la méthode effectivement utilisée)

Cette méthode est restituée ici par la formule matricielle `computeLongitudesAllMethodsLongitude(obliquité de l’écliptique, longitudes des 6 maisons, marge d’erreur, nombre de lignes)`.
Elle produit quatre tableaux : longitudes théoriques, coefficients de qualité, ascensions droites et intervalle de latitude géographique.

### Implémentation Excel

Bien qu'il soit possible d'utiliser ces méthodes globales grâce à des formules matricielles, l'avantage d'Excel tient plutôt à la possibilité de les reconstituer en utilisant des formules intermédiaires.

La feuille de calcul d'example propose donc trois onglets :
* **ARRAYS** présente un formulaire permettant recourant aux formules matricielles pour les deux méthodes (de la même façon que la feuille de calcul modèle de _[Horosc for Google Sheets](https://github.com/Turbehrt/horosc-GoogleSheets)_)
* **METHOD A** détaille les calculs successifs pour la méthode A.
  - Le tableau du haut convertit les arguments des dégrés sexagésimaux en radians, calcule chaque longitude et ascension droite en radians, puis reconvertit ces résultats en degrés.
  - Le second tableau réalise les mêmes calculs directement en degrés (sans afficher les résultats intermédiaires en radians).
* **METHOD B** détaille de la même façon
  - en haut à gauche : les longitudes et ascensions droites théoriques et les coefficients de qualité (distance entre les longitudes théoriques et observées) en radians
  - en bas à gauche : les mêmes en degrés
  - en haut à droite : la croix des latitudes en radians et en degrés

À la différence des formules matricielles de l'onglet ARRAYS, chaque cellule de METHOD A et METHOD B donne le résultat d'une formule distincte. Ces formules peuvent êtres réutilisées dans d'autres cellules, ou importées dans d'autres projets Excel en y copiant les trois modules VBA.

> [!NOTE]
> Dans les formulaires proposés (comme dans le programme Horosc initial), tous les nombres en entrée sont exprimés en degrés sous forme sexagésimale. Si les séparateurs sont cohérents, ils sont réutilisés dans les résultats (ex : 187.12'04, 187°12'04'', 187d 12m). Cependant, il est également possible d'utiliser les formules pour calculer à partir de nombres décimaux en radians.
>
> Lors de la comparaison avec des données issues du programme Horosc initial, ou de _[Horosc for Google Sheets](https://github.com/Turbehrt/horosc-GoogleSheets)_, notez que les longitudes, ascensions droites et latitudes géographiques fournies en résultat y sont exprimées en degrés sexagésimaux, mais les coefficients de qualité sont en radians.

## Méthodes de calcul et fonctions intermédiaires

### Séquences de calcul

Les deux méthodes enchaînent les calculs comme suit :
* Méthode A (`computeLongitudesAllMethodsLatitude`)
  - conversion des entrées en radians
  - calcul de l’ascension droite de l’ascendant
  - calcul de l’ascension droite du fond du ciel (en utilisant la différence ascensionnelle)
  - calcul de la longitude du fond du ciel
  - appel de chaque méthode et conversion des résultats en degrés (sexagésimaux)
  - affichage des résultats

* Méthode B (`computeLongitudesAllMethodsLongitude`)
  - conversion des entrées en radians
  - extraction de la longitude de l’ascendant et du fond du ciel
  - calcul des ascensions droites de l’ascendant et du fond du ciel
  - calcul de la latitude théorique du lieu d’observation
  - appel de chaque méthode et conversion des résultats en degrés (sexagésimaux)
  - calcul des coefficients de qualité (en radians)
  - conversion de la marge d’erreurs en radians et calcul d’un intervalle de latitudes :
    + au centre la valeur théorique (d’après l’ascendant et le fond du ciel fournis)
    + verticalement, les valeurs en cas d’erreur/approximation de l’ascension droite du fond du ciel
    + horizontalement, les valeurs en cas d’erreur/approximation de l’ascension droite de l’ascendant
  - affichage des résultats

### Fonctions intermédiaires


> [!IMPORTANT]
> Toutes les formules personnalisées sont codées sous formes de fonctions dans trois modules VBA interdépendants : **Trigonometry** (conversions de base), **Domification** (calculs individuels) and **Sequences** (fonctions globales enchaînant les calculs).

* **Calcul d'angles** (module _Trigonometry_)
  + `ModuloRange`, `ModuloTwoPI`
  + `SplitSexagesimalFormat`, `SexagesimalFormat` : extrait les nombres et séparateurs utilisés en entrée dans une chaîne de caractères représentant un nombre sexagésimal
  +	`SexagesimalToRadian` : convertit un nombre sexagésimal en radians
  +	`RadianToSexagesimal` : convertit un nombre en radians en degrés, en utilisant un ensemble de séparateurs (facultatif)

* **Coordonnées célestes** (module _Domification_)
  +	`EclipticToEquator(obliquity, longitude)` : convertit une longitude en ascension droite (étant connue l’obliquité de l’écliptique)
  +	`EquatorToEcliptic(obliquity, rightAscension)` : convertit une ascension droite en longitude (étant connue l’obliquité de l’écliptique)
  +	`RetrieveLatitude(obliquity, rightASC, rightIMC)` (radians), `RetrieveLatitudeSexagesimal` (degrés) : calcule la latitude théorique du lieu d’observation à partir de l’obliquité de l’écliptique, et des ascensions droites de l’ascendant et du fond du ciel (IMC).
  +	`RetrieveLatitudeFromLong(obliquity, longASC, longIMC)` (radians), `RetrieveLatitudeFromLongSexagesimal` (degrés) : calcule la latitude théorique du lieu d’observation à partir de l’obliquité de l’écliptique, et des longitudes de l’ascendant et du fond du ciel (IMC).
  +	`RetrieveLatitudeRange(obliquity, longASC, longIMC, error, direction)` (radians), `RetrieveLatitudeRangeSexagesimal` (degrés) : application d'une marge d'erreur (`error`) dans le calcul de la latitude, en suivant la croix proposée par North
    * `direction = 0` : aucune erreur (identique à `RetrieveLatitude`)
    * `direction = 1` (gauche) ou `direction = 2` (droite) : marge d'erreur appliquée à l'ascension droite de l'ascendant
    * `direction = 3` (haut) ou `direction = 4` (bas) : marge d'erreur appliquée à l'ascension droite du fond du ciel

> [!IMPORTANT]
> La formule `RetrieveLatitudeRange` corrige des incohérences constatées dans le code PASCAL du programme initial. Elle ne retourne donc pas les mêmes résultats que le programme en PASCAL ou que la version 1 de *Horosc for Google Sheets* (mais elle est cohérente avec la version 2 de *Horosc for Google Sheets*). Voir la section [Différences avec le programme initial de J. D. North](#diff%C3%A9rences-avec-le-programme-initial-de-j-d-north) pour plus de détails, et sur la façon de restituer la formule initiale en cas de besoin.

* **Domification**  (module _Domification_) : pour chaque méthode de domification, les fonctions `Method0(obliquity, geoLatitude, rightASC, rightIMC, houseIndex, getRA)` à `Method6` renvoient l'ascension droite (`getRA = true`) ou la longitude (`getRA = false`) de la pointe d'une maison (`houseIndex` de 1 à 6) à partir de l'obliquité de l'écliptique (`obliquity`), de la latitude géographique du lieu d’observation (`geoLatitude`) et de l'ascension droite de l'ascendant et du fond du ciel (`rightASC`, `right IMC`). Entrées et résultats en radians.
  + `Method0` : méthode des lignes horaires (_Hour Lines method, fixed boundaries_). Les pointes sont les intersections de l'écliptique avec l'horizon, le cercle méridien et les lignes des heures inégales (paires). Cette méthode est généralement graphique, à l'aide d'un astrolabe, ici émulé avec une fonction de convergence (`Converge`).
  + `Method1` : méthode Standard, dite d'Alcabitius (_Standard method_). Division uniforme des secteurs cardinaux de l'équateur.
  + `Method2` : méthode à double longitude (_Dual longitude method_). Division uniforme des secteurs cardinaux de l'écliptique.
  + `Method3` : méthode du Premier Vertical (_Prime Vertical method, fixed boundaries_). Division uniforme du Premier Vertical.
  + `Method4` : méthode équatoriale (_Equatorial method, fixed boundaries_). Division uniforme de l'équateur sur la sphère locale.
  + `Method5` : méthode équatoriale à limites mobiles (_Equatorial method, moving boundaries_). Division uniforme de l'équateur sur la sphère céleste.
  + `Method6` : méthode à longitude simple (_Single Longitude Method_). Division uniforme de l'écliptique.

> [!NOTE]
> Pour l'explication détaillée et des cas d'usages historiques de chaque méthode, voir John D. North, *Horoscopes and History*, London: Warburg Institute, 1986.


* **Coefficients de qualité** (module _Domification_)
  + `QualityCoefficientRadian(observedLongitude, computedLongitude)`, `QualityCoefficientDegree` : les coefficients de qualité correspondent à la différence entre une longitude observée (transcrite d'une source historique) et une longitude calculée.

> [!NOTE]
> Dans le programme Horosc initial, les coefficients de qualité sont exprimés en radians, quand toutes les autres valeurs sont en degrés.


* **Fonctions globales** (module _Sequences_), permettant d'enchaîner les conversions en fonction des entrées disponibles
  + `computeCuspWithMethodInRadian(obliquity, geoLatitude, rightASC, longASC, rightIMC, longIMC, houseIndex, method, getRA)` : calcule les coordonnées (ascension droite ou longitude) de la pointe de n'importe quelle maison (1-6), selon n'importe quelle méthode (0-6), à partir de l'ensemble des paramètres connus : obliquité de l'écliptique, latitude géographique du lieu d'observation, ascensions droites et longitudes de l'ascendant et du fond du ciel (en radians)
  + `ComputeCuspFromLatitudeInRadian(obliquity, geoLatitude, longASC, houseIndex, method, getRA)` : calcule les coordonnées (ascension droite ou longitude) de la pointe de n'importe quelle maison (1-6), selon n'importe quelle méthode (0-6), à partir de la longitude de l'ascendant et de la latitude géographique d'observation (en radians) -- sans connaître les coordonnées du fond du ciel
  + `computeCuspFromLatitudeInSexagesimal` :  calcule les coordonnées (ascension droite ou longitude) de la pointe de n'importe quelle maison (1-6), selon n'importe quelle méthode (0-6), à partir de la longitude de l'ascendant et de la latitude géographique d'observation (en degrés)
  + `computeCuspFromLongitudeInRadian(obliquity, longASC, longIMC, houseIndex, method, getRA)` : calcule les coordonnées (ascension droite ou longitude) de la pointe de n'importe quelle maison (1-6), selon n'importe quelle méthode (0-6), à partir des longitudes de l'ascendant et du fond du ciel (en radians) -- sans connaître les ascensions droites ou la latitude d'observation
  + `computeCuspFromLongitudeInSexagesimal` : calcule les coordonnées (ascension droite ou longitude) de la pointe de n'importe quelle maison (1-6), selon n'importe quelle méthode (0-6), à partir des longitudes de l'ascendant et du fond du ciel (en degrés)

* **Fonctions matricielles** (module _Sequences_)
  + `ComputeLongitudesAllMethodsLatitude`: [méthode A](#methode-a) (voir plus haut)
  + `computeLongitudesAllMethodsLongitude`: [méthode B](#methode-b) (voir plus haut)
  + fonctions annexes pour présenter les résultats : `moveRows`, `fillEmpty`

### Différences avec le programme initial de J. D. North

Plusieurs choix algorithmiques du programme en Pascal, en particulier pour la restitution des latitudes, ont paru surprenants et n'ont pas été transposés à l'identique.

* le calcul de la latitude théorique du lieu d'observation suppose que l'observation se fait toujours dans l'hémisphère Nord en appliquant une valeur absolue. **Ce comportement est préservé par défaut dans le présent programme.** A titre expérimental, un argument facultatif `methNorth` a été introduit dans la fonction `RetrieveLatitude` : si `methNorth = false`, la latitude est modulée entre $-\pi/2$ et $\pi/2$ radians.

> [!WARNING]
> Cet algorithme n'a pas encore été testé ; il n'est pas recommandé d'y recourir pour l'instant (risque de résultat faux et/ou incohérent avec d'autres fonctions).

* la composition initiale de la croix des latitudes nous a paru incohérente. Les fonctions `FOI` et `FIO` du code Pascal, correspondant à nos directions 1 et 2 de `RetrieveLatitudeRange` (marge d'erreur appliquée à l'ascendant) retiraient également la marge d'erreur à l'ascension droite du fond du ciel, ce qui n'est pas cohérent avec la logique de calcul. **La correction est appliquée par défaut depuis la version 1 du présent programme.**

| `RetrieveLatitude` | | |
| --- | --- | --- |
| | (**3**) `RightASC`, `RightIMC + error` |  |
| (**1**) `RightASC - error`, `RightIMC` | (**0**) `RightASC`, `RightIMC` | (**2**) `RightASC + error`, `RightIMC` |
| :x: _(**FOI**) `RightASC - error`, `RightIMC - error`_ | (**4**) `RightASC`, `RightIMC - error` | :x: _(**FIO**) `RightASC + error`, `RightIMC + error`_ | 

> [!TIP]
> Il reste cependant possible de calculer manuellement FOI : `RetrieveLatitude(obliquity, rightASC - error, rightIMC - error)` et FIO : `RetrieveLatitude(obliquity, rightASC + error, rightIMC - error)`.

* les marges d'erreur utilisées pour le calcul de la croix des latitudes étaient associées en Pascal à une valeur en radians ne correspondant pas à leur libellé. **Les marges d'erreur annoncées (en degrés) sont utilisées dans le présent programme** au lieu des anciennes valeurs en radians, ce qui peut peut amener à des estimations de latitudes différentes de celles du programme initial.
  + _\[D] 1 min. arc_	était associé à 0.001745329 rad, en réalité	0°06'00"
  + _\[E] half min._ était associé à 0.00087266 rad, en réalité	0°03'00"
  + _\[F] 1 sec. arc_ était associé à 0.000029089 rad, en réalité 0°00'06"
