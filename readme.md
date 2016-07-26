RcppXLSXwriter
================

[![Project Status: WIP ? Initial development is in progress, but there has not yet been a stable, usable release suitable for the public.](http://www.repostatus.org/badges/latest/wip.svg)](http://www.repostatus.org/#wip) [![Travis-CI Build Status](https://travis-ci.org/carlganz/RcppXLSXwriter.svg?branch=master)](https://travis-ci.org/carlganz/RcppXLSXwriter)

This package uses John McNamara's C library [libxlsxwriter](http://libxlsxwriter.github.io/) and [Rcpp](http://www.rcpp.org/) to let R users generate customized XLSX files in R. The [`WriteXLS`](https://cran.r-project.org/web/packages/WriteXLS/index.html) package on CRAN is nice, but it requires Perl, and it isn't flexible enough to allow the user to fully customize their Excel output.

With `RcppXLSXwriter` you have the ability to customize each and every cell in your XLSX document. \#\# Issues

### Portability

So far I have only been able to get the package to build on 32-bit R on Windows.
