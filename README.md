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

The program written by John D. North offered two approaches:

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

## Calculation methods and intermediate functions

### Sequences

The two methods perform the calculations as follows:

* Method A (`computeLongitudesAllMethodsLatitude`)
  - conversion of inputs to radians
  - calculation of the right ascension of the ascendant
  - calculation of the right ascension of the *Imum Coeli* (using the ascensional difference)
  - calculation of the longitude of the *Imum Coeli*
  - calling each method and converting the results to degrees (sexagesimal)
  - display of results

* Method B (`computeLongitudesAllMethodsLongitude`)
  - conversion of inputs to radians
  - extraction of the longitude of the ascendant and *Imum Coeli*
  - calculation of the right ascensions of the ascendant and *Imum Coeli*
  - calculation of the theoretical latitude of the observation site
  - calling each method and converting the results to (sexagesimal) degrees
  - calculation of quality coefficients (in radians)
  - conversion of the margin of error into radians and calculation of a latitude interval:
    + at the centre, the theoretical value (based on the ascendant and *Imum Coeli* provided)
    + vertically, the values in case of error/approximation of the right ascension of the ascendant
    + horizontally, the values in case of error/approximation of the right ascension of the *Imum Coeli*
  - display of results

### Intermediate functions

> [!IMPORTANT]
> All custom formulas are coded in three interdependent VBA modules: **Trigonometry** (basic sexagesimal conversions), **Domification** (individual calculations) and **Sequences** (global functions chaining individual opeartions).

* **Angle calculation** (from _Trigonometry_)
  + `ModuloRange`, `ModuloTwoPI`
  + `SplitSexagesimalFormat`, `SexagesimalFormat`: extracts the individual numbers and separators used in the input string
  +  `SexagesimalToRadian`: converts a sexagesimal number to radians
  +  `RadianToSexagesimal`: converts a number in radians to degrees, using a set of separators (optional)

* **Celestial Coordinates** (from _Domification_)
  + `EclipticToEquator(obliquity, longitude)`: converts ecliptic longitude to right ascension (given the obliquity of the ecliptic).
  + `EquatorToEcliptic(obliquity, rightAscension)`: converts right ascension to ecliptic longitude (given the obliquity of the ecliptic).
  + `RetrieveLatitude(obliquity, rightASC, rightIMC)` (radians), `RetrieveLatitudeSexagesimal` (degrees): calculates the theoretical latitude of the observation site based on the obliquity of the ecliptic and the right ascensions of the Ascendant and *Imum Coeli* (IMC).
  + `RetrieveLatitudeFromLong(obliquity, longASC, longIMC)` (radians), `RetrieveLatitudeFromLongSexagesimal` (degrees): calculates the theoretical latitude based on the obliquity and the ecliptic longitudes of the Ascendant and *Imum Coeli* (IMC).
  + `RetrieveLatitudeRange(obliquity, longASC, longIMC, error, direction)` (radians), `RetrieveLatitudeRangeSexagesimal` (degrees): applies an error margin (`error`) to the latitude calculation, following the "cross" method proposed by North.
    * `direction = 0`: no error (identical to `retrieveLatitude`).
    * `direction = 1` (up) or `direction = 2` (down): error margin applied to the Ascendant's longitude.
    * `direction = 3` (left) or `direction = 4` (right): error margin applied to the *Imum Coeli*'s longitude.

> [!IMPORTANT]
> The `RetrieveLatitudeRange` formula fixes inconsistencies found in the original PASCAL code. Consequently, it does not return the same results as the PASCAL program or version 1 of *Horosc for Google Sheets*, but it is consistent with version 2 of *Horosc for Google Sheets*. See [Differences with J.D. North's original program](#differences-with-jd-norths-original-program) for more details, anf how to emulate the original formula.

* **House Division**  (from _Domification_): for each calculation method, functions `Method0(obliquity, geoLatitude, rightASC, rightIMC, houseIndex, getRA)` through `Method6` return either the right ascension (`getRA = true`) or the longitude (`getRA = false`) of a house cusp (`houseIndex` from 1 to 6). Inputs and results are in radians.
  + `Method0`: _Hour Lines method_ (fixed boundaries). Cusps are the intersections of the ecliptic with the horizon, the meridian, and the unequal (even) hour lines. This method is traditionally graphical (using an astrolabe), emulated here via a convergence function (`Converge`).
  + `Method1`: _Standard method_, known as Alcabitius. Uniform division of the equatorial cardinal sectors.
  + `Method2`: _Dual longitude method_. Uniform division of the ecliptic cardinal sectors.
  + `Method3`: _Prime Vertical method_ (fixed boundaries). Uniform division of the Prime Vertical.
  + `Method4`: _Equatorial method_ (fixed boundaries). Uniform division of the equator on the local sphere.
  + `Method5`: _Equatorial method_ (moving boundaries). Uniform division of the equator on the celestial sphere.
  + `Method6`: _Single Longitude method_. Uniform division of the ecliptic.

> [!NOTE]
> Detailed explanation and historical use of each method are to be found by John D. North, *Horoscopes and History*, London: Warburg Institute, 1986.

* **Quality Coefficients**
  + `QualityCoefficientRadian(observedLongitude, computedLongitude)`, `QualityCoefficientDegree`: these represent the difference between an observed longitude (as transcribed from a historical source) and a calculated longitude.

> [!NOTE]
> Note that in the original North programme, quality coefficients are expressed in radians, while all other quantities are in sexagesimal degrees.


* **Global functions** (from _Sequences_): these allow for chained conversions based on available inputs.
  + `computeCuspWithMethodInRadian(obliquity, geoLatitude, rightASC, longASC, rightIMC, longIMC, houseIndex, method, getRA)`: computes the coordinates (right ascension or longitude) of the cusp of any house (1-6), based on any method (0-6), using all known parameters: obliquity, geographical latitude, right ascensions and longitudes of the Ascendant and _Immum Coeli_ (in radians)
  + `computeCuspFromLatitudeInRadian(obliquity, geoLatitude, longASC, houseIndex, method, getRA)`: computes the coordinates (right ascension or longitude) of the cusp of any house (1-6), based on any method (0-6), using the Ascendant's longitude and the geographic latitude (in radians) -- without knowing the coordinates of the _Immum Coeli_.
  + `computeCuspFromLatitudeInSexagesimal`: computes the coordinates (right ascension or longitude) of the cusp of any house (1-6), based on any method (0-6), using the Ascendant's longitude and the geographic latitude (in degrees).
  + `computeCuspFromLongitudeInRadian(obliquity, longASC, longIMC, houseIndex, method, getRA)`: computes the coordinates (right ascension or longitude) of the cusp of any house (1-6), based on any method (0-6), using the longitudes of the Ascendant and _Immum Coeli_ (in radians).
  + `computeCuspFromLongitudeInSexagesimal`: computes the coordinates (right ascension or longitude) of the cusp of any house (1-6), based on any method (0-6), using the longitudes of the Ascendant and _Immum Coeli_ (in degrees).

* **Array functions** (from _Sequences_):
  + `ComputeLongitudesAllMethodsLatitude`: [method A](#method-a) (see above)
  + `computeLongitudesAllMethodsLongitude`: [method B](#method-b) (see above)


### Differences with J. D. North's Original Program

Several algorithmic choices in the original Pascal program, particularly regarding latitude retrieval, appeared surprising and were not reproduced identically.

* The calculation of the theoretical latitude assumes that observations always take place in the Northern Hemisphere by applying an absolute value. **This behavior is preserved by default in this program.** Experimentally, an optional `methNorth` argument has been introduced in the `retrieveLatitude` function: if `methNorth = false`, the latitude is modulated between $-\pi/2$ and $\pi/2$ radians. This algorithm has not yet been fully tested; using it is currently not recommended (risk of incorrect or inconsistent results).

* The original composition of the "latitude cross" appeared inconsistent. The `FOI` and `FIO` functions in the Pascal code -- corresponding to our directions 1 and 2 in `RetrieveLatitudeRange` (error margin applied to the Ascendant) -- also subtracted the error margin from the right ascension of the _Immum Coeli_, which is inconsistent with the calculation logic. **This correction has been applied by default since version 1.** It remains possible to manually calculate FOI as `RetrieveLatitude(obliquity, rightASC - error, rightIMC - error)` and FIO as `RetrieveLatitude(obliquity, rightASC + error, rightIMC - error)`.

* The margins of error used to calculate the "latitude cross" were associated in Pascal with a value in radians that did not actually correspond to their label. **The margins of error announced (in degrees) are used in this program** instead of the old values in radians, which may lead to latitude estimates that differ from those in the initial program).
  + \[D] "1 min. arc" was associated with 0.001745329 rad, in reality 0°06'00"
  + \[E] "half min." was associated with 0.00087266 rad, in reality 0°03'00"
  + \[F] "1 sec. arc" was associated with 0.000029089 rad, in reality 0°00'06"
