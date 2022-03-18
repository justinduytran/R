\name{extractABSTableBuilder}
\alias{extractABSTableBuilder}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
%%  ~~function to do ... ~~
Extract data from ABS TableBuilder exports.
}
\description{
%%  ~~ A concise (1-5 lines) description of what the function does. ~~
This function reads the awkwardly formatted Australian Bureau of Statistic's (ABS) TableBuilder exports and formats it into something more parsable by other software by transforming it into a long format data set.

}
\usage{
extractABSTableBuilder(path, wafer = TRUE, totals = FALSE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
%%     ~~Describe \code{x} here~~
  \item{path}{Path to .xls or .xlsx export from TableBuilder. Does not support other formats.
}
  \item{wafer}{Designate whether wafers are present in the export or not. Function by default assumes yes (TRUE).
}
  \item{totals}{Indicate whether to include the 'Totals' supplied by the ABS. Default is no (FALSE).
}

}
\details{
%%  ~~ If necessary, more details than the description above ~~
The function works by reading the first sheet and finding the summation 'Total' row and column cell references. These references are used to define the sheet area to loop over the sheets in the export and extract data. Column names and other identifiers are (mostly) derived relative to this sheet area.

Relies on the 'tidyverse' package and 'readxl' from tidyverse (also installs pacman to load these packages).
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
%##---- Should be DIRECTLY executable !! ----
%##-- ==>  Define data, use random,
%##--	or do  help(data=index)  for the standard data sets.

% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory (show via RShowDoc("KEYWORDS")):
% \keyword{ ~kwd1 }
% \keyword{ ~kwd2 }
% Use only one keyword per line.
% For non-standard keywords, use \concept instead of \keyword:
% \concept{ ~cpt1 }
% \concept{ ~cpt2 }
% Use only one concept per line.
