# horosc-Excel

Porting John D. North's historical domification program Horosc to Microsoft Excel (VBA)

## Proposition for two methods of computation

### Method A : Compute cusp of houses from known points, using different historical methods

Providing you already have:

- Obliquity
- Geographical Latitude
- Longitude of ascendant

The function `computeCuspFromLatitude` computes the longitudes and right ascensions for a specific house (1 to 6) and a specific method (0 to 6)

### Method B: Evaluate the accuracy of observed longitudes

providing you already have:

- Obliquity
- Observed longitudes of the houses

The function `retrieveLatitudeFromLong` will return the geographical Latitude
the function `RetrieveLatitudeRange` will provide insights on variation of the Latitude when applying an Error RTange on either ascendant or IMC
The function `computeCuspFromLongitude` will computes the theorical longitudes and right ascensions for a specific house (1 to 6) and a specific method (0 to 6). It will consider observed longitudes for house 1 and 4 as correct.
You can then use the `QualityCoefficient` function to get the comparison between observed longitude and theorical longitude

## The 7 methods for computing cusp of houses

## Convention names

- Obliquity is the obliquity of the ecliptic
- geoLat is the geographical latitude of the place (usually phi), equivalent to the elevation of the poles (over the horizon)
- ascensionalDifference is the ascensional difference (right ascendant - right ascension of the vernal point)

Any domification methods computes the cusps of each House (1 to 6, with houses 7 to 12 being symetrical).They come as right Ascensions (on the equinox, usually noted alpha) or longitudes (on the ecliptic, usually noted lambda).

Meaningful cusps are the Ascendant (Asc, cusp of House 1, intersection of ecliptic and horizon) and the Immum Caeli (IMC, cusp of House 4, intersection of ecliptic and night meridian)

right ascension of the ascendant - ascensional difference = right ascension of the IMC - 90°
Houses 7 to 12 are symetrical to Houses  to 6 in all methods (cusp + 180)

## Implementation of the computations in EXCEL VBA

Using a spreadsheet to layout a presentation of results of computations imply that the hroscope computations must be provided under the form of funtions.
We built a set of function that make easy to build a presentation or what you want to show.

The core part of these functions are:

- the 7 historical methods of computing the cusp of houses
- the logic to compute geographical latitude based on ascendant and IMC, although this function could be replaced by a formula using standard functions of XL
these core functions are implemented using angles in radian.

Other functions are only helper functions

- Basic trigonometry transformations
- Convenient way to invoke the core functions with transformation of parametres or call historical methods based on simple index.

## Examples provided on the spreadsheet

## In the details of the code provided

### Sexagesimal, Radians and Degrees

### Available functions

| **module** | **content** |
| :--- | :--- |
|||

#### Core module : Domification

There are 7 methods (numbered 0 to 6) to compute the cusp of house.
Each method take input parameters:

- Obliquity, geographic Latitude: values expected in radian
- One or several of: Right Ascendant, Long Ascendant, right IMC, long IMC - values expected in radian
- the houseIndex indicating which cusp to compute - value from 1 to 6
- an optional bolean to indicate waht value return:
  - FALSE (or not provided) - returm the cusp right Ascendion in radian
  - TRUE - return the cusp longitude in radian

| **function** | **parameters** | **description and usage** |
| :---: | :-------- | :-------- |
|**Method0**<br><br>**Hour Lines<br>(fixed boundaries)**|- obliquity<br>- geographic Latitude<br>- right Ascendant<br>- right IMC<br>- houseIndex<br>- RA or Long|Cusps are intersections of the ecliptic by the horizon, the meridian circle, and the unequal hour lines on the sphere for even-numbered hours<br><br>This method is usually graphical, with aid of an astrolabe.<br>It is emulated using a convergence function|
|**Method1**<br><br>**Standard method (Alcabitius)**|- obliquity<br>- geographic Latitude<br>- right Ascendant<br>- right IMC<br>- houseIndex<br>- RA or Long|Uniform division of of the cardinal sectors of the Equator|
|**Method2**<br><br>**Dual longitude**|- obliquity<br>- geographic Latitude<br>- long Ascendant<br>- long IMC<br>- houseIndex<br>- RA or Long|Uniform division of cardinal sectors of the Ecliptic|
|**Method3**<br><br>**Prime Vertical (fixed boundaries)**|- obliquity<br>- geographic Latitude<br>- right IMC<br>- long IMC<br>- houseIndex<br>- RA or Long|Uniform division of the Prime Vertical|
|**Method4**<br><br>**Equatorial (fixed boundaries)**|- obliquity<br>- geographic Latitude<br>- right IMC<br>- long IMC<br>- houseIndex<br>- RA or Long|Uniform division of the Equator (local sphere)|
|**Method5**<br><br>**Equatorial (moving boundaries)**|- obliquity<br>- geographic Latitude<br>- right Ascendant<br>- houseIndex<br>- RA or Long|Uniform division of the Equator (celestial sphere)|
|**Method6**<br><br>**Single longitudes**|- obliquity<br>- geographic Latitude<br>- long Ascendant<br>- houseIndex<br>- RA or Long|Uniform division of the Ecliptic|

Some other useful function for conversion of coordinate:

| **function** | **parameters**| **description and usage** |
| :--- | :--- | :--- |
|**EclipticToEquator**|- Obliquity<br>- longitutde| transform the longitude in right ascension |
|**EquatorToEcliptic**|- Obliquity<br>- right ascension| transform the right ascension into a longitude|
|**retrieveLatitude**|- Obliquity<br>- right Ascendant<br>- right IMC|compute the geographical Latitude based on Ascendant and IMC positions|

#### Helper module : radian conversion functions

| **function** | **parameters**| **description and usage** |
| :--- | :--- | :--- |
|**SexagesimalToRadian**|- sexa<br>- fmt (optional)| Expect to have the parameter sexa filled with an angle in sexagesimal notation<br>Exemple: 182°32'10 or 182d32'10 or 182.32.10 or 182 32 10. Separator of the 3 elements can be, by default, one of : °, (space), (dot), (quote), (double-quote), d. Caller can provide a different set of chars as separators.<br>example: ```SexagesimalToRadian("180-00-00", "-") -> Pi value```|
|**RadianToSexagesimal**|- radian<br>- fmt (optional)| Transform the radian value provided into a string sexagesimal representation. It usee by default degree sign and quote as separtor of degrees and minutes<br>Example: ```RadianToSexagesimal(pi/2) -> "90°00'00"```.<br><br>The caller can provide its own char to separate the sexgesimal parts.<br>Example: ```RadianToSexagesimal(pi/2, "=-") => "90=00-00"```  |

#### Facade module : convenient set of functions that invole the core functions in a more user-friendly way

## Adaptation from initial Pascal implementation
