\name{extractABS}
\alias{extractABS}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Extract data from ABS tables with multiple data sheets%%  ~~function to do ... ~~
}
\description{
This function merges together all the 'Data' sheets in a typical ABS table into one data.table in R.
It removes the need to manually merge each Data1, Data2,...,DataN sheet together.
Please see the details section for assumptions and limitations.
The function relies on the 'readxl' and 'datatable' packages.%%  ~~ A concise (1-5 lines) description of what the function does. ~~
}
\usage{
extractABS(filepath)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{filepath}{This should point to the .xls or .xlsx file downloaded from the ABS
%%     ~~Describe \code{filepath} here~~
}
}
\details{
Assumes the 'Data' sheets in the excel file are consecutive, i.e. assumed ABS sheet format is: Index, Data1,Data2,...,DataX, Inquiries, Misc, etc.
It assumes there are 9 column identifiers i.e. Unit Type, Series Type, Data Type, Data Interval, etc and removes them.
%%  ~~ If necessary, more details than the description above ~~
}
\value{
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
%%  ~~who you are~~
}
\note{
%%  ~~further notes~~
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\seealso{
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
\examples{
##---- Should be DIRECTLY executable !! ----
##-- ==>  Define data, use random,
##--	or do  help(data=index)  for the standard data sets.

## The function is currently defined as
function (filepath)
{
    library("readxl")
    library("data.table")
    all_sheets = excel_sheets(filepath)
    data_sheets = as.character(all_sheets[(grep("Data", all_sheets))])
    excel_data <- lapply(data_sheets, read_excel, path = filepath)
    names(excel_data) <- data_sheets
    n = length(excel_data)
    ABS_extract <- cbind(excel_data$Data1, eval(parse(text = paste("cbind(",
        paste("excel_data$Data", 2:n, "[,-1]", sep = "", collapse = ", "),
        ")"))))
    ABS_extract = as.data.table(ABS_extract)
    ABS_extract = ABS_extract[-1:-9, ]
    setnames(ABS_extract, 1, "Date", skip_absent = TRUE)
    ABS_extract = as.data.table(lapply(ABS_extract, as.numeric))
    ABS_extract[, 1] = as.Date(unlist(ABS_extract[, 1]), origin = "1900-01-01")
    ABS_extract
  }
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory (show via RShowDoc("KEYWORDS")):
% \keyword{ ~kwd1 }
% \keyword{ ~kwd2 }
% Use only one keyword per line.
% For non-standard keywords, use \concept instead of \keyword:
% \concept{ ~cpt1 }
% \concept{ ~cpt2 }
% Use only one concept per line.
