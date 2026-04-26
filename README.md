# Horosc for Excel

[![en](https://img.shields.io/badge/lang-en-red.svg)](/README.md)
[![fr](https://img.shields.io/badge/lang-fr-blue.svg)](/README.fr.md)

## Project background

John D. North's book, *Horoscopes and History*, London: Warburg Institute, 1986, published in an appendix (appendix 4, pp. 197-218) the Pascal for MS-DOS code of an application entitled **HOROSC** and designed to calculate and check the domification of a horoscope using the seven main historical methods.

Due to the obsolescence of the Pascal language, this application became very difficult to access. Launched in 2021, the project _[Horosc for Google Sheets](https://github.com/Turbehrt/horosc-GoogleSheets)_ ported the code to Google Script for use in Google Sheets.

The way Google Apps Script works favors the execution of complete methods (not unlike the initial Horosc program) instead of a more flexible use of intermediate functions. This project, launched in 2025, ports the code to VBA for Microsoft Excel, creating a wider range of functions to be combined in different and more customized way.

> [!TIP]
> Since version 2 of _[Horosc for Google Sheets](https://github.com/Turbehrt/horosc-GoogleSheets)_, both applications use the same functions and arguments names, so as to facilitate parallel use. Running the Excel spreadsheet on a local computer tends to be faster than loading cloud-based Google Apps scripts, especially when calling several intermediate functions. However, this requires macros to be enabled.

## General principle

The programme written by John D. North offered two approaches:

### Method A

Knowing:
*    the value of the obliquity of the ecliptic,
*    the geographical latitude of the observation location,
*    and the ecliptic longitude of the Ascendant

the application provides:

* the theoretical calculation of the sexagesimal longitudes and/or right ascensions of the first 6 houses (the next 6 are inferred by symmetry) according to the 7 historical house systems

This method is reproduced here by the array formula `computeLongitudesAllMethodsLatitude(obliquity of the ecliptic, geographical latitude, longitude of the ascendant, number of lines)`.
It produces two tables, one for longitudes and one for right ascensions, separated by the number of lines entered.

### Method B

Knowing:
*    the value of the obliquity of the ecliptic,
*    the longitudes of the first 6 houses (transcribed from a historical source),
*    and a margin of error or rounding

the application provides:
*    a theoretical calculation of the latitude of the observation site (with an interval corresponding to the margin of error, applied to the right ascension of the ascendant or midheaven)
*    a comparison with the theoretical longitudes (calculated considering only the ascendant and the midheaven) according to the seven historical house systems, with a quality coefficient (generally allowing the method actually used to be identified)

This method is reproduced here by the array formula `computeLongitudesAllMethodsLongitude(obliquity of the ecliptic, longitudes of the 6 houses, margin of error, number of lines)`.
It produces four tables: theoretical longitudes, quality coefficients, right ascensions and geographical latitude interval.

### Excel implementation

Although it is possible to use these global methods through array formulas in an Excel spreadsheet, the advantage rather lies in reconstituting them by using intermediate formulas.

The sample spreadsheet thus provides three tabs:
* **ARRAYS** offers a template to use both methods as array formulas (similar to the model spreadsheet of _[Horosc for Google Sheets](https://github.com/Turbehrt/horosc-GoogleSheets)_)
* **METHOD A** breaks down the calculations for Method A. The top table converts the arguments from sexagesimal degrees to radians, calculates each longitude and right ascension in radians, and then converts them back to sexagesimal degrees. The second table performs the same calculations directly in degrees (the conversions are not shown).
* **METHOD B** similarly computes
  - the theoretical longitudes, rigth ascensions and quality coefficients (distance between observed and theoretical longitudes) in radians in the top left corner
  - the same in sexagesimal degrees in the bottom left corner
  - the latitude cross in radians and degrees the top right corner

Unlike in ARRAYS, each cell of METHOD A and METHOD B is the result of a single formula. Those fomulas can be used in other cells, or imported to other Excel projects by copying the three VBA modules.

> [!NOTE]
> In the proposed templates (as in the original Horosc programme), all input numbers are expected to be expressed in degrees in sexagesimal form. Separators provided in the input, if consistant, are reused in the output (ex: 187.12'04, 187°12'04'', 187d 12m). However, it is possible to use formulas to compute from inputs in decimal radians.
> 
> When comparing with data from the original North programme, or with _[Horosc for Google Sheets](https://github.com/Turbehrt/horosc-GoogleSheets)_, note that resulting longitudes, right ascensions and geographical latitudes are expected to be expressed in sexagesimal degrees, but quality coefficients in radians.




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
