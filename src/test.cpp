#include <Rcpp.h>
#include <xlsxwriter.h>
using namespace Rcpp;

//' @useDynLib libxlsxwriter
//' @export
//' @import Rcpp methods
// [[Rcpp::export]]
void helloworld(std::string file_,const CharacterMatrix& M,const CharacterVector& t) {

  int n=M.nrow();int m=M.ncol();
  const char *file=file_.c_str();
  lxw_workbook  *workbook  = workbook_new(file);
  lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

  lxw_format *formatB = workbook_add_format(workbook);
  format_set_bold(formatB);

  for (int i=0;i<m;i++) {
    const char *name=std::string(t(i)).c_str();
    worksheet_write_string(worksheet, 0, i,name, formatB);
  }

  for (int i=0;i<m;i++) {
    for (int j=0;j<n;j++) {
      const char *temp=std::string(M(j,i)).c_str();
      worksheet_write_string(worksheet, j+1, i,temp, NULL);
    }
  }

  workbook_close(workbook);

}
