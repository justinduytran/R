# ------------------------------------------------------------------------------
# Function to extract from an Australian Bureau of Statistics (ABS) Table Builder export
# ------------------------------------------------------------------------------
# Requires:
# path - specify path to .xlsx (or whatever format) Table Builder export file
# wafer - boolean variable indicating if the extract contains wafers, default assumes yes (TRUE)
# totals - indicate whether the Totals supplied by the ABS are required, default is no as it stuffs with Excel Pivot Tables

extractABSTableBuilder = function(path, wafer = TRUE, totals = FALSE ){
  if (!require("pacman")) install.packages("pacman")
  pacman::p_load(tidyverse,readxl)
  
  # Find the cells to be extracted from each sheet. This is done by finding two corner points as defined by the 'Total' row and column
  raw_extract=read_excel(path)
  bounds=which(raw_extract == "Total", arr.ind = TRUE)
  
  # Reduce these bounds if totals are not required
  if (isTRUE(totals)) {
    removetotals=0
  } else {
    removetotals=-1
  }
  
  # Shamelessly stolen from https://stackoverflow.com/a/52214227
  # This functions converts column reference numbers to excel column references (i.e. 2 -> B, 27 -> AA)
  # It also goes the other way (i.e. C -> 3. X -> 24) but we don't need that
  # This is needed later to convert bounds to cell ranges for our read_excel function
  xlcolconv <- function(col){
    # test: 1 = A, 26 = Z, 27 = AA, 703 = AAA
    if (is.character(col)) {
      # codes from https://stackoverflow.com/a/34537691/2292993
      s = col
      # Uppercase
      s_upper <- toupper(s)
      # Convert string to a vector of single letters
      s_split <- unlist(strsplit(s_upper, split=""))
      # Convert each letter to the corresponding number
      s_number <- sapply(s_split, function(x) {which(LETTERS == x)})
      # Derive the numeric value associated with each letter
      numbers <- 26^((length(s_number)-1):0)
      # Calculate the column number
      column_number <- sum(s_number * numbers)
      return(column_number)
    } else {
      n = col
      letters = ''
      while (n > 0) {
        r = (n - 1) %% 26  # remainder
        letters = paste0(intToUtf8(r + utf8ToInt('A')), letters) # ascii
        n = (n - 1) %/% 26 # quotient
      }
      return(letters)
    }
  }
  
  # Create blank table that we can append our 'for loop' to
  tidy_extract=data.frame()
  
  # Count the number of sheets we need to use our 'for loop' for
  sheets=excel_sheets(path) %>%
    # Remove the 'template_rse' and 'format' sheets that ABS has hidden
    head(-2)
  
  # For Loop to iterate over every 'wafer' sheet
  for (i in sheets){
    # This section uses the previously extracted bounds to pull out our cell ranges
    temp_extract=read_excel(path,
                            range = 
                              paste(xlcolconv(bounds[1,2]),bounds[2,1]+2,
                                    ":",
                                    xlcolconv(bounds[2,2]+removetotals),bounds[1,1]+2+removetotals,
                                    sep=""),
                            sheet = i
    )
    # Remove first row and fix labeling (blame ABS for their crappy layout)
    colnames(temp_extract)[1] = c(temp_extract[1,1])
    temp_extract <- temp_extract[-c(1),]
    
    # This reads the sheet's 'wafer' (if specified that wafers exist) and adds it as a defining column
    if (isTRUE(wafer))
      {wafer_data=read_excel(path,
                             range = paste(xlcolconv(bounds[1,2]-1),
                                           bounds[2,1]+2-1,
                                           sep=""),
                             sheet = i) %>%
        colnames
      
    # Add wafer identifier to the wafer
    temp_extract=cbind(wafer_data,temp_extract)
    }
  
  # Append this sheet to our final data set
  tidy_extract=rbind(tidy_extract,temp_extract)
  
  #End loop
  }

  # Correctly label first column
  wafer_name=raw_extract[1,1] %>%
    str_extract(".*?by") %>%
    str_sub(end=-3) %>%
    str_trim()
  colnames(tidy_extract)[1] = wafer_name
  
  # Make into 'Long' data set (so Excel can read it into a Pivot Table)
  if (isTRUE(wafer)){
    tablebuilderextract=pivot_longer(tidy_extract,cols = -c(1,2))
  } else {
    tablebuilderextract=pivot_longer(tidy_extract,cols = -c(1))
  }
  
  # Correctly label TableBuilder 'column variable' column
  colnames(tablebuilderextract)[length(names(tablebuilderextract))-1] = raw_extract[bounds[2,1],1]
  
  # Rename 'value' column to show ABS Counting method
  colnames(tablebuilderextract)[length(names(tablebuilderextract))] = raw_extract[2,1]
  
  # Remove 'Total' wafer if totals' are not required
  if (isFALSE(totals)) {
  tablebuilderextract=filter(tablebuilderextract, tablebuilderextract[,1] != "Total")
  }
  
  # Print some notifiers for QoL
  print("Extraction from Table Builder table complete")
  print("Please use write_csv if you wish to export and use this data in an Excel Pivot Table")

  # Function finished, return our happy data set :)
  return(tablebuilderextract)
}

