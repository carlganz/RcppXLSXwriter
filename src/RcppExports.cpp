// This file was generated by Rcpp::compileAttributes
// Generator token: 10BE3573-1514-4C36-9D1C-5A225CD40393

#include <Rcpp.h>

using namespace Rcpp;

// helloworld
void helloworld(std::string file_, const CharacterMatrix& M, const CharacterVector& t);
RcppExport SEXP RcppXLSXwriter_helloworld(SEXP file_SEXP, SEXP MSEXP, SEXP tSEXP) {
BEGIN_RCPP
    Rcpp::RNGScope __rngScope;
    Rcpp::traits::input_parameter< std::string >::type file_(file_SEXP);
    Rcpp::traits::input_parameter< const CharacterMatrix& >::type M(MSEXP);
    Rcpp::traits::input_parameter< const CharacterVector& >::type t(tSEXP);
    helloworld(file_, M, t);
    return R_NilValue;
END_RCPP
}
