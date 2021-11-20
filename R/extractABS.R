#test changes
extractABS=function(filepath,id_rows=9){
  #required libraries
  library("readxl")
  library("data.table")
  #pull file name from function input
  #give list of excel sheet names
  all_sheets=excel_sheets(filepath)
  #find excel sheets with 'Data' in title
  data_sheets=as.character(all_sheets[(grep("Data",all_sheets))])
  #ignore sheets without 'Data' in title
  excel_data <- lapply(data_sheets, read_excel, path = filepath)
  #rename the sheets to their actual names (DataX) instead of generic names
  names(excel_data) <- data_sheets
  #find out how many Data sheets there are
  n=length(excel_data)
  #merge all the sheets into one data frame
  ABS_extract <-
    #all of the 'Data1' sheet
    cbind(excel_data$Data1,
          #the rest of the DataX sheets,
          #eval & parse for the function to interpret the concatenated input
          eval(parse(text=paste("cbind(",
                                #reiterate merge from Data2 to DataN
                                paste("excel_data$Data", 2:n,
                                      #ignore the first (date) column for each
                                      #unneeded as it is already taken from Data1
                                      "[,-1]", sep = "", collapse = ", "), ")"))))
  #convert to data.table
  ABS_extract=as.data.table(ABS_extract)
  #remove the rows at the start with unneeded data (column identifiers)
  ABS_extract=ABS_extract[-1:-id_rows,]
  #name the first column - Date
  setnames(ABS_extract,1,"Date",skip_absent = TRUE)
  #convert the table from characters into numeric values
  ABS_extract=as.data.table(lapply(ABS_extract, as.numeric))
  #convert the date column from excel dates to R dates
  ABS_extract[,1]=as.Date(unlist(ABS_extract[,1]),origin="1900-01-01")
  #output the data.table
  ABS_extract
}