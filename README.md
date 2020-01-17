# MS Excel utility functions

This repository contains MS Excel spreadsheets in .xlsx format. Whether or not they work in OpenOffice, LibreOffice, or alike is unknown to me. Due to Excels limitations they are not by far as advanced as comparable R code, and some of the used Excel equations are rather complicated and confusing.

All codes in this repository are distributed under the [Creative Commons Attribution-NonCommercial-ShareAlike 3.0 Unported License](http://creativecommons.org/licenses/by-nc-sa/3.0/). For further information, please visit my [website](https://sites.google.com/view/manuel-f-g-weinkauf/home).

# Description of functions

## Alkalinity.xlsx v. 1.0

This Excel spreadsheet allows to calculate the alkalinity of seawater based on pH measurements obtained from titrating seawater with HCl.

## BenjaminiHochberg.xlsx v. 1.1

This Excel spreadsheet provides means to calculate corrected levels of significance for cases of multiple testing. Both the Benjamini and Hochberg (1995) correction and the simple Bonferroni correction are calculated. The only limitation, in comparison to the R version, is, that no more than 50 *p*-values can be entered without manually changing some formulas—but this will apply seldom if ever, anyway.

* Benjamini, Y. and Hochberg, Y. (1995) Controlling the false discovery rate: A practical and powerful approach to multiple testing. *Journal of the Royal Statistical Society B* 57 (1): 289–300.

## Conductivity2Salinity.xlsx v. 1.0

This Excel sheet allows to convert conductivity values of sea water (under known temperature and pressure) into salinity values corresponding to the practical salinity unit scale.

## CoordinatesTranslator.xlsx v. 1.0

An early and very basic version of the R code also provided. It was prepared to convert large amounts of degrees–minute–second coordinates into decimal degrees for purposes of program comparability. Nothing special, nothing complicated, nothing one could not easily do by oneself; but now its there and someone may have use for it.

## MultinomialConfidence.xlsx v. 1.0

The calculation of the confidence interval of relative abundances of species, morphotypes, etc. in a sample is not as easy as in many other cases. Since relative abundances cannot be negative and individual relative abundances of several groups in a sample are not independent of each other, procedures for multinomial mathematics must be applied. Luckily, Heslop et al. (2011) developed a modern adaptation of those methods. The spreadsheet provided reads relative abundances and calculates the multinomial confidence intervals for all species.

* Heslop, D., De Schepper, S., and Proske, U. (2011) Diagnosing the uncertainty of taxa relative abundances derived from count data. *Marine Micropaleontology* 79: 114–20.

## PAST2_Template.xlsx v. 1.0

Allows to comfortably edit data in MS Excel and export them as a .dat file (including row- and column-names and symbol-colour-encoding per row) understood by PAST v. 2.x.

## PAST3_Template.xlsx v. 1.0

Allows to comfortably edit data in MS Excel and export them as a .dat file (including row- and column-names, column types, and symbol- and colour-encoding per row) understood by PAST v. 3.x.

## PseudoLandmarks.xlsx v. 1.0

When landmarks which describe morphological structures as outlines are extracted by hand, individual landmarks are naturally not equally spaced along that structure-outline. Since equal spacing is often a necessity for morphometric analysis, this Excel sheet allows to convert such a set of landmarks into pseudo-landmarks which are equally spaced along the reconstructed outline of the structure.

## StereoMean.xlsx v. 2.1

This spreadsheet was created some years ago for the purpose of calculating the mean of several structural geological data collected in the field, together with some statistical tests for the sufficiency of sample size of those data. Though not completely free of minor problems, I provide it here because it does a reasonable job in more than 90% of cases (and also because I am not likely to ever improve it further, anyway, so possibly someone else will do the job).

# Author

Manuel F. G. Weinkauf, Université de Genève, Departement of Earth Sciences, 1205 Geneva, Switzerland; [Manuel.Weinkauf@unige.ch](mailto:Manuel.Weinkauf@unige.ch); [https://sites.google.com/site/weinkaufmanuel/](https://sites.google.com/site/weinkaufmanuel/); [http://www.unige.ch/sciences/terre/people/personal_pages/ManuelWeinkauf/ManuelWeinkauf.php](http://www.unige.ch/sciences/terre/people/personal_pages/ManuelWeinkauf/ManuelWeinkauf.php)

















