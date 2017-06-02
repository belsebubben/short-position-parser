# short-pos-parser
works with stocks on the swedish stock market. Will retreive and parse a stock search by name and see the short posisitons held by companies and how they are acting on those positions.

## about
short positions above 0.5% of company have to be reported in sweden to finansinspektionen (fi.se) so information is public. However many holders have positions and it can be a bit daunting to get a good picture of the totals and how institutions are positioning themselves

## requirements
a unix system
requirements python3 
xlrd==1.0.0
xlwt==1.2.0

## usage example
### part of the company name like (skf or volv should be enough)
short-position-parser atlas
Retreiving:  http://www.fi.se/contentassets/71a61417bb4c49c0a4a3a2582ea8af6c/korta_positioner_2017-06-01.xlsx
could not convert string to float:  ; ;;;
'Holding' ; ;;;
'Holding' ; ;;;
could not convert string to float:  ; ;INGA PUBLIKA POSITIONER PUBLICERADES/NO PUBLIC POSITIONS WERE PUBLISHED;;

Results

Nr:  1
Name: ATLAS COPCO	Isin: SE0000101032	Total: 0.56%
LANSDOWNE:                     Holding:  0.56 Increasing:  -3

Nr:  2
Name: ATLAS COPCO	Isin: SE0006886750	Total: 0.83%
LANSDOWNE:                     Holding:  0.34 Increasing:  0
AQR:                           Holding:  0.49 Increasing:  -1

Merged Names ['ATLAS COPCO', 'ATLAS COPCO']
Merged Isins ['SE0000101032', 'SE0006886750']
Merged Result
Name: ATLAS COPCO	Isin: SE0000101032	Total: 1.61%
LANSDOWNE:                     Holding:  0.56 Increasing:  -4
AQR:                           Holding:  0.49 Increasing:  -1

