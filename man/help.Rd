\name{docinfeR}
\alias{docinfeR}
\title{Inference Doc Report Generation with R}
\description{Generation of Inference Report Analysis in .docx format as appears in Statsframe ULTRA software (thank you to Statsframe <https://statsframe.com> team for providing excerpts of the developed code for this package)}
\usage{docinfeR(path,var.type)}
\arguments{
	\item{path}{the path to a .csv file for analysis and subsequential report output}
	\item{var.type}{the type of variables to perform analysis, with possible values: "all", "numeric", "integer", "double", "factor", "character"}
}
\value{A file in .docx with the output report}
\author{Rui Portocarrero Sarmento, Vera Costa}
\examples{\dontrun{docinfeR(path="data.csv",var.type="numeric")}}
\keyword{inference}

