# ******************************************************************************
# Created: 26-Aug-2024
# Last modified: 26-Aug-2024
# Contributor(s): httk
# ******************************************************************************


#' Save dictionary for cleaning value labels and variable labels
#'
#' This function allows users to generate a Excel file which includes variable labels
#' and value labels for reviewing and cleaning.
#'
#' @param x datainput (dataframe).
#' @param filename name of the output Excel file to be saved.
#' @param data_sheet (optional) state sheet name if would like to export the data.
#' @param dict_sheet state sheet name of dictionary sheet
#' @param ... Additional arguments may apply.
#' @return Working dictionary of the datafile
#' @examples
#' # Set the example folder of the package as the working directory
#' setwd(system.file("extdata", package = "epinionPS"))
#'
#' # Read the data
#' mydata <- epinion_read_data(file = "sample_test_report.sav")
#'
#' # Save dictionary for cleaning value labels and variable labels
#' epinion_write_labelled_xlsx(mydata, "dic.xlsx", data_sheet = "")

# Source: expss::write_labelled_xlsx
epinion_write_labelled_xlsx = function(x,
                                       filename,
                                       data_sheet = "data",
                                       dict_sheet = "dictionary",
                                       remove_repeated = FALSE,
                                       use_references = FALSE){
  if(!requireNamespace("openxlsx", quietly = TRUE)){
    stop("write_labelled_xlsx: 'openxlsx' is required for this function. Please, install it with 'install.packages('openxlsx')'.")
  }
  stopifnot(
    is.data.frame(x),
    length(filename)==1L,
    is.character(filename),
    length(remove_repeated)==1L,
    remove_repeated %in% c(TRUE, FALSE),
    length(use_references)==1L,
    use_references %in% c(TRUE, FALSE),
    is.character(dict_sheet),
    length(dict_sheet) == 1L
  )
  wb = openxlsx::createWorkbook()
  if(data_sheet == ""){
    # do nothing
  }
  else
  {
    sh = openxlsx::addWorksheet(wb, sheetName = data_sheet)
    openxlsx::writeData(wb = wb,
                        sheet = sh,
                        x = unlab(x),
                        borderColour = "black",
                        borderStyle = "none",
                        keepNA = FALSE)
    openxlsx::freezePane(wb, sh, firstCol = TRUE, firstRow = TRUE)
  }
  dict = epinion_create_dictionary(x,
                                   remove_repeated = remove_repeated,
                                   use_references = use_references
  )

  dict$label <- lapply(dict$label, epinion_cleaning_tool)

  if(nrow(dict)>0){
    sh = openxlsx::addWorksheet(wb, sheetName = dict_sheet)
    openxlsx::writeData(wb = wb,
                        sheet = sh,
                        x = dict,
                        borderColour = "black",
                        borderStyle = "none",
                        keepNA = FALSE)
    openxlsx::freezePane(wb, sh, firstCol = TRUE, firstRow = TRUE)
    openxlsx::addStyle(wb, sh, style = createStyle(textDecoration = "bold"), rows = 1, cols = 1:6)
  }
  openxlsx::saveWorkbook(wb, filename, overwrite = TRUE)
}

# ==============================================================================
#' Read cleaned dictionary and apply to the data
#'
#' This function allows users to apply cleaned variable labels and value labels to
#' the dataframe.
#'
#' @param x datainput (dataframe).
#' @param filename name of the input dictionary (Excel file).
#' @param dict_sheet state sheet name of dictionary sheet
#' @param ... Additional arguments may apply.
#' @return A dataframe in which variable labels and value labels are cleaned
#' @examples
#' # Set the example folder of the package as the working directory
#' setwd(system.file("extdata", package = "epinionPS"))
#'
#' # Read the data
#' mydata <- epinion_read_data(file = "sample_test_report.sav")
#'
#' # Save dictionary for cleaning value labels and variable labels
#' epinion_write_labelled_xlsx(mydata, "dic.xlsx", data_sheet = "")
#'
#' # Read cleaned dictionary and apply to the data
#' mydata <- epinion_read_labelled_xlsx(mydata, "dic - Cleaned.xlsx")

# Source: expss::read_labelled_xlsx
epinion_read_labelled_xlsx = function(x,
                                      filename,
                                      dict_sheet = "dictionary"){
  if(!requireNamespace("openxlsx", quietly = TRUE)){
    stop("read_labelled_xlsx: 'openxlsx' is required for this function. Please, install it with 'install.packages('openxlsx')'.")
  }
  stopifnot(
    length(filename)==1,
    is.character(filename),
    length(dict_sheet)==1,
    is.numeric(dict_sheet) || is.character(dict_sheet)
  )
  wb = openxlsx::loadWorkbook(file = filename)

  sheet_names = names(wb)
  if((dict_sheet %in% sheet_names) ||(dict_sheet %in% seq_along(sheet_names))){
    dict = openxlsx::readWorkbook(wb,
                                  sheet = dict_sheet,
                                  colNames = TRUE,
                                  rowNames = FALSE,
                                  skipEmptyRows = FALSE,
                                  check.names = FALSE,
                                  na.strings = "NA"
    )
    x = epinion_apply_dictionary(x, dict)
  } else {
    if(!missing(dict_sheet)){
      message("read_labelled_xlsx: sheet '", dict_sheet,
              "' with dictionary not found. Labels will not be applied to data.")
    }
  }
  x
}

# ==============================================================================
# Source: expss::apply_dictionary
epinion_apply_dictionary = function(x, dict){
  stopifnot(is.data.frame(x),
            is.data.frame(dict),
            all(c("variable", "value", "label", "meta") %in% colnames(dict))
  )
  if(nrow(dict)==0) return(x)
  dict[["variable"]][dict[["variable"]] %in% ""] = NA
  dict[["meta"]][dict[["meta"]] %in% ""] = NA
  # fill NA
  for(i in seq_len(nrow(dict))[-1]){
    if(is.na(dict[["variable"]][i])) {
      dict[["variable"]][i] = dict[["variable"]][i - 1]
    }
  }
  dict = dict[dict$variable %in% colnames(x), ]
  # variable labels
  all_varlabs = dict[dict$meta %in% "varlab",]
  truncated_varlabs = all_varlabs[nchar(as.character(all_varlabs$label)) >= 256,]

  vallabs = dict[dict$meta %in% NA,]

  for(i in seq_len(nrow(all_varlabs))){
    if(!is.na(all_varlabs$label[i])) {
      var_label(x[[all_varlabs$variable[i]]]) = all_varlabs$label[i]
    } else {
      var_label(x[[all_varlabs$variable[i]]]) = ""
    }

  }
  # value labels
  references = dict[dict$meta %in% "reference",]
  missing_references = setdiff(unique(references$label), names(vallabs))
  if(length(missing_references)>0){
    warning(paste0(" missing references - ", paste(paste0("'", missing_references, "'"), collapse = ", ")))
    references = references[references$label %in% names(vallabs), ]
  }
  for(i in seq_len(nrow(references))){
    val_label(x[[references$variable[i]]], references$value[i]) = references$label[i]
  }

  categorical_vars = dict[dict$meta %in% NA,]
  truncated_vallabel = categorical_vars[nchar(as.character(categorical_vars$label)) >= 120,]

  for(i in seq_len(nrow(categorical_vars))){
    if (class(x[[categorical_vars$variable[i]]])[1] != "factor") {
      val_labels(x[[categorical_vars$variable[i]]]) <- NULL
    }
  }

  for(i in seq_len(nrow(categorical_vars))){
    if (class(x[[categorical_vars$variable[i]]])[1] != "factor") {
      val_label(x[[categorical_vars$variable[i]]], categorical_vars$value[i]) = categorical_vars$label[i]
    }
  }

  truncated_infor = rbind(truncated_varlabs, truncated_vallabel)

  if(nrow(truncated_infor)>0){
    wb = openxlsx::createWorkbook()
    sh = openxlsx::addWorksheet(wb, sheetName = "Truncated infor")
    openxlsx::writeData(wb = wb,
                        sheet = sh,
                        x = truncated_infor,
                        borderColour = "black",
                        borderStyle = "none",
                        keepNA = FALSE)
    openxlsx::freezePane(wb, sh, firstCol = TRUE, firstRow = TRUE)
    openxlsx::saveWorkbook(wb, "Truncated infor.xlsx", overwrite = TRUE)
  }

  x
}

# ==============================================================================
epinion_cleaning_tool = function(x){
  if(!requireNamespace("textclean", quietly = TRUE)){
    pacman::p_load(textclean)
  }

  # Replace Repeated Whitespace with a Single Space
  x = gsub(" +"," ", x)

  # Trim Leading and Trailing Whitespace
  x = trimws(x)

  # Removes \n!
  x = gsub("\\s+"," ",x)

  # Replace a new line with a space
  x = gsub("\\n"," ",x)

  # Remove HTML/XML tags (basic)
  x = rm_angle(x)

  # Remove some special text at Epinion surveys
  redundant_text <- c(" (Sæt gerne flere kryds)", " (sæt gerne flere kryds)",
                      " (Mulighed for flere svar)", "'' : @", " : @", " : resp", "' ",
                      " (mulighed for flere kryds)", "Flere svar mulig",
                      " -- Flere svar mulig.", " Angiv antal",
                      " (Vælg alle svar, der passer)--", " (Vælg alle svar, der passer)",
                      " -- Flere svar mulig.", "Flere svar mulig")

  # x = qdap::mgsub(redundant_text, "",x)

  x = textclean::mgsub(x, redundant_text, "")

  # Remove "'" at the beginning and at the end
  x = gsub("^'","",x)
  x = gsub("\\'$","",x)

  # Remove ":" at the end
  x = gsub("\\:$","",x)

  # Clean "&amp;" to "&"
  x = gsub("&amp;","&",x)

  x
}
