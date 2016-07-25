helloworld <- function(df, file) {
  title <- names(df)
  d <- dim(df)
  M <- as.character(as.matrix(df))
  dim(M) <- d
  invisible(.Call('libxlsxwriter_helloworld', PACKAGE = 'libxlsxwriter', file, M, title))
}
