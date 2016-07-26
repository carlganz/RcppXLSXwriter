#include <Rcpp.h>
#include <xlsxwriter.h>
using namespace Rcpp;

RCPP_EXPOSED_CLASS(XLSXworkbook);
RCPP_EXPOSED_CLASS(XLSXworksheet);
RCPP_EXPOSED_CLASS(XLSXformat);

class XLSXworkbook {
public:
  XLSXworkbook(std::string file_):file(file_) {
    const char *filec=file.c_str();
    workbook=workbook_new(filec);
  }

  lxw_workbook *workbook;
  std::string file;
};

void closewb(XLSXworkbook *wb) {
  workbook_close(wb->workbook);
}



class XLSXworksheet {
public:
  XLSXworksheet(XLSXworkbook workbook_,std::string sheet_):workbook(workbook_),
  sheet(sheet_) {
    const char *sheetc=sheet.c_str();
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

void writef(XLSXworksheet *ws,IntegerVector x,IntegerVector y,CharacterMatrix value,XLSXformat *format) {
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

void write(XLSXworksheet *ws,IntegerVector x,IntegerVector y,CharacterMatrix value) {
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

void font_color(XLSXformat *format,lxw_color_t color) {
  format_set_font_color(format->format,color);
}

void background_color(XLSXformat *format,lxw_color_t color) {
  format_set_bg_color(format->format,color);
}

void foreground_color(XLSXformat *format,lxw_color_t color) {
  format_set_fg_color(format->format,color);
}

void font_size(XLSXformat *format, int size) {
  format_set_font_size(format->format,size);
}

void font_name(XLSXformat *format, std::string font_) {
  const char *font=font_.c_str();
  format_set_font_name(format->format,font);
}

void bold(XLSXformat *format) {
  format_set_bold(format->format);
}

void italic(XLSXformat *format) {
  format_set_italic(format->format);
}

void underline(XLSXformat *format) {
  format_set_underline(format->format, LXW_UNDERLINE_SINGLE);
}

RCPP_MODULE(workbook_mod) {
  class_<XLSXworkbook>("XLSXworkbook")

  .constructor<std::string>()
  .method("close",&closewb,"Creates XLSX file")
  .field_readonly("file",&XLSXworkbook::file,"File name")
;
  class_<XLSXworksheet>("XLSXworksheet")
    .constructor<XLSXworkbook,std::string>()
    .field_readonly("sheet",&XLSXworksheet::sheet,"Sheet name")
    .method("write",&write,"Write to XLSX file without formatting. Use a matrix.")
    .method("writef",&writef,"Write to XLSX file with formatting. Use a matrix.")
  ;
  class_<XLSXformat>("XLSXformat")
    .constructor<XLSXworkbook>()
    .method("bold",&bold,"Bold formatting")
    .method("italic",&italic,"Italic formatting")
    .method("underline",&underline,"Underline formatting")
    .method("font_color",&font_color,"Font color formatting")
    .method("background_color",&background_color,"Background color formatting")
    .method("foreground_color",&foreground_color,"Foreground color formatting")
    .method("font_name",&font_name,"Font name formatting")
    .method("font_size",&font_size,"Font size formatting")
  ;
}


