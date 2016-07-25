#include <Rcpp.h>
#include <xlsxwriter.h>
using namespace Rcpp;

RCPP_EXPOSED_CLASS(XLSXworkbook);
RCPP_EXPOSED_CLASS(XLSXworksheet);
RCPP_EXPOSED_CLASS(XLSXformat);

class XLSXworkbook {
public:
  XLSXworkbook(std::string file_):file(file_) {
    const char * filec=file.c_str();
    workbook=workbook_new(filec);
  }

  lxw_workbook *workbook;
  std::string file;
};

void closewb(XLSXworkbook* wb) {
  workbook_close(wb->workbook);
}



class XLSXworksheet {
public:
  XLSXworksheet(XLSXworkbook workbook_,std::string sheet_):workbook(workbook_),
  sheet(sheet_) {
    const char * sheetc=sheet.c_str();
    worksheet=workbook_add_worksheet(workbook.workbook,sheetc);
  }

  XLSXworkbook workbook;
  std::string sheet;
  lxw_worksheet *worksheet;

};


class XLSXformat {
public:
  XLSXformat(XLSXworkbook workbook_):
  workbook(workbook_) {
    format=workbook_add_format(workbook.workbook);
  }
  lxw_format *format;
  XLSXworkbook workbook;
};

void writef(XLSXworksheet* ws,IntegerVector x,IntegerVector y,CharacterMatrix value,XLSXformat* format) {
  x=x-1;y=y-1;
  for (int i=0;i<x.size();i++) {
    for (int j=0;j<y.size();j++) {
      int xv = x(i);int yv = y(j);
      std::string val=Rcpp::as<std::string>(value(i,j));
      const char* cval=val.c_str();
      worksheet_write_string(ws->worksheet,xv,yv,cval,format->format);
    }
  }
}

void write(XLSXworksheet* ws,IntegerVector x,IntegerVector y,CharacterMatrix value) {
  x=x-1;y=y-1;
  for (int i=0;i<x.size();i++) {
    for (int j=0;j<y.size();j++) {
      int xv = x(i);int yv = y(j);
      std::string val=Rcpp::as<std::string>(value(i,j));
      const char* cval=val.c_str();
      worksheet_write_string(ws->worksheet,xv,yv,cval,NULL);
    }
  }
}

void bold(XLSXformat *format) {
  format_set_bold(format->format);
}

void red(XLSXformat *format) {
  format_set_font_color(format->format,LXW_COLOR_RED);
}

void green(XLSXformat *format) {
  format_set_font_color(format->format,LXW_COLOR_GREEN);
}

void italic(XLSXformat *format) {
  format_set_italic(format->format);
}

void underline(XLSXformat*format) {
  format_set_underline(format->format, LXW_UNDERLINE_SINGLE);
}

RCPP_MODULE(workbook_mod) {
  class_<XLSXworkbook>("XLSXworkbook")

  .constructor<std::string>()
  .method("close",&closewb)
  .field_readonly("file",&XLSXworkbook::file)
;
  class_<XLSXworksheet>("XLSXworksheet")
    .constructor<XLSXworkbook,std::string>()
    .field_readonly("sheet",&XLSXworksheet::sheet)
    .method("write",&write)
    .method("writef",&writef)
  ;
  class_<XLSXformat>("XLSXformat")
    .constructor<XLSXworkbook>()
    .method("bold",&bold)
    .method("red",&red)
    .method("green",&green)
    .method("italic",&italic)
    .method("underline",&underline)
  ;
}


